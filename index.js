const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const cheerio = require("cheerio");
const axios = require("axios");
const puppeteer = require("puppeteer-core");
const chromium = require('@sparticuz/chromium');
require("dotenv").config();

const app = express();
const PORT = 3000;

app.use(bodyParser.json());

let serviceAccount;

try {
  serviceAccount = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
} catch (err) {
  console.error("❌ Failed to load service account JSON:", err.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

app.get("/test", async (req, res) => {
  const data = "section#tickets";
  try {
    const browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
    });
    const page = await browser.newPage();
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
      "AppleWebKit/537.36 (KHTML, like Gecko) " +
      "Chrome/125.0.0.0 Safari/537.36"
    );


    await page.goto("https://www.ticketmaster.co.uk/download/terms-and-conditions", {
      waitUntil: "domcontentloaded",
    });


    const pageTitle = await page.title();


    const row = [];
    let targetElement = `${data} div.container`;


    const content = await page.$$eval(targetElement, (sections) => {
      return sections.map((section) => {
        const h3 = section.querySelector("h3")?.innerText.trim() || section.querySelector("h2")?.innerText.trim();
        const items = Array.from(section.querySelectorAll(".find-ticket-items"))
          .map((item) => {
            const h4 = item.querySelector("h4")?.innerText.trim() || "";
            const p = item.querySelector("p")?.innerText.trim() || "";
            const href = item.querySelector("a")?.href || "";
            return [h4, p, href ? `=HYPERLINK("${href}", "Click here")` : ""];
          });
        return { title: h3, items };
      });
    });


    content.forEach(section => {
      row.push([section.title, "Date", "Link"]);
      section.items.forEach(item => row.push(item));
      row.push(["", "", ""]);
    });


    await browser.close();


    res.status(200).json({ row });
  } catch (err) {
    console.error(err);
    res.status(403).json({ message: err.message });
  }
});

app.post("/data", async (req, res) => {
  const { sheet, link, content } = req.body;

  if (!sheet || !link || !content) {
    return res
      .status(400)
      .json({ error: "Sheet name, link and content are required." });
  }

  try {
    // const browser = await puppeteer.launch({});
    const browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: chromium.headless,
    });
    const page = await browser.newPage();
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
      "AppleWebKit/537.36 (KHTML, like Gecko) " +
      "Chrome/125.0.0.0 Safari/537.36"
    );

    await page.goto(link, { waitUntil: "domcontentloaded" });

    await clearSheetDataAndFormatting(sheet);

    const values = await parseContent(content, page, link);

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: sheet,
      valueInputOption: "USER_ENTERED",
      resource: { values },
    });

    await formatSheet(sheet, values);

    await browser.close();

    res.status(200).json({ message: "Data saved", values });
  } catch (err) {
    console.error("❌ Error saving data:", err);
    res.status(500).json({ error: "Failed to save data." });
  }
});

async function clearSheetDataAndFormatting(sheetName) {
  const sheetId = await getSheetId(sheetName);

  // 1. Clear data
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName,
  });

  // 2. Clear formatting
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: {
              sheetId,
            },
            fields: "userEnteredFormat",
          },
        },
      ],
    },
  });

  console.log(`✅ Cleared data and formatting in "${sheetName}"`);
}

async function parseContent(data, page, link) {
  const rows = [];
  const pageTitle = await page.title();
  rows.push([pageTitle, link, '']);
  rows.push(['', '', '']);


  // Try the 'tickets' selector first, then fallback to 'ticket-container'
  let targetElement = `${data} div.container div.tickets`;
  const hasTickets = await page.$(targetElement);
  if (!hasTickets) {
    targetElement = `${data} div.container div.ticket-container`;
  }


  // Wait a bit for the targetElement to appear (helpful for SPA / JS-rendered pages)
  try {
    await page.waitForSelector(targetElement, { timeout: 5000 });
  } catch (e) {
    // not fatal — we'll continue and $$eval will return [] if nothing matches
  }


  // We pass no external variables except what's needed (none here) and do detection inside the page
  const sections = await page.$$eval(targetElement, (sections) => {
    return sections.map((section) => {
      const h3 = (section.querySelector('h3') && section.querySelector('h3').innerText.trim()) ||
        (section.querySelector('h2') && section.querySelector('h2').innerText.trim()) || '';


      // prefer .find-ticket-items; if .ticket-row exists within this section, use that instead
      let itemSelector = '.find-ticket-items';
      if (section.querySelector('.ticket-row')) itemSelector = '.ticket-row';


      const items = Array.from(section.querySelectorAll(itemSelector)).map((item) => {
        const h4 = item.querySelector('h4') ? item.querySelector('h4').innerText.trim() : '';
        const p = item.querySelector('p') ? item.querySelector('p').innerText.trim() : '';
        const a = item.querySelector('a');
        const href = a ? a.href : '';
        return [h4, p, href ? `=HYPERLINK("${href}", "Click here")` : ''];
      });


      return { title: h3, items };
    });
  });


  sections.forEach((s) => {
    rows.push([s.title, 'Date', 'Link']);
    s.items.forEach((it) => rows.push(it));
    rows.push(['', '', '']);
  });


  return rows;
}

async function formatSheet(sheet, values) {
  const sheetId = await getSheetId(sheet);

  const rowCount = values.length;
  const colCount = Math.max(...values.map((row) => row.length));

  const requests = [];

  // Detect header rows (with 'Date' and 'Link')
  const headerRows = values
    .map((row, i) => {
      if (
        row[1] === "Date" &&
        row[2] === "Link" &&
        typeof row[0] === "string"
      ) {
        return i;
      }
      return null;
    })
    .filter((i) => i !== null);

  // Center all cells
  requests.push({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: 0,
        endRowIndex: rowCount,
        startColumnIndex: 0,
        endColumnIndex: colCount,
      },
      cell: {
        userEnteredFormat: {
          textFormat: { fontSize: 12 },
          horizontalAlignment: "CENTER",
          verticalAlignment: "MIDDLE",
          wrapStrategy: "WRAP",
        },
      },
      fields:
        "userEnteredFormat(horizontalAlignment,wrapStrategy,verticalAlignment)",
    },
  });
  requests.push({
    autoResizeDimensions: {
      dimensions: {
        sheetId, // get this via Sheets API or from metadata
        dimension: "COLUMNS",
        startIndex: 0, // e.g. start from column A
        endIndex: 3, // exclusive: up to column C
      },
    },
  });

  // Bold + large headers
  headerRows.forEach((rowIndex) => {
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: rowIndex,
          endRowIndex: rowIndex + 1,
          startColumnIndex: 0,
          endColumnIndex: colCount,
        },
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, fontSize: 14 },
            horizontalAlignment: "CENTER",
            wrapStrategy: "WRAP",
          },
        },
        fields:
          "userEnteredFormat(wrapStrategy,textFormat,horizontalAlignment)",
      },
    });
  });

  // Add borders row by row, skipping blank rows
  values.forEach((row, rowIndex) => {
    const isBlank = row.every((cell) => !cell || cell.trim() === "");
    if (isBlank) return;

    requests.push({
      updateBorders: {
        range: {
          sheetId,
          startRowIndex: rowIndex,
          endRowIndex: rowIndex + 1,
          startColumnIndex: 0,
          endColumnIndex: colCount,
        },
        top: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
        bottom: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
        left: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
        right: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
        innerHorizontal: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
        innerVertical: {
          style: "SOLID",
          width: 1,
          color: { red: 0, green: 0, blue: 0 },
        },
      },
    });
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
}

async function getSheetId(sheetName) {
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const sheet = metadata.data.sheets.find(
    (s) => s.properties.title === sheetName,
  );

  return sheet.properties.sheetId;
}

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
