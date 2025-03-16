const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const cheerio = require("cheerio");
const axios = require("axios");
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
  const data = "section#terms";
  try {
    const respo = await axios.get(
      "https://www.ticketmaster.co.uk/trnsmt/terms",
    );
    const $ = cheerio.load(respo.data);
    const pageTitle = $("title").text().trim();
    const row = [];
    $(`${data} div.container div.tickets`).each((_, section) => {
      const $section = $(section);
      const h3Text = $section.find("h3").first().text().trim();
      row.push([h3Text, "date", "link"]);

      // Only get h4, p, and a inside .find-ticket-items
      $section.find(".find-ticket-items").each((_, item) => {
        const h4 = $(item).find("h4").first().text().trim();
        const p = $(item).find("p").first().text().trim();
        const rawHref = $(item).find("a").first().attr("href");
        const href = rawHref ? `=HYPERLINK("${rawHref}", "Click here")` : "";

        row.push([h4, p, href]);
      });

      row.push(["", "", ""]); // Optional spacer
    });
    console.log(row);

    res.status(200).json({ message: pageTitle });
  } catch (err) {
    res.status(400).json({ message: err });
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
    const response = await axios.get(link);
    const html = response.data;

    await clearSheetDataAndFormatting(sheet);

    const values = parseContent(content, html, link);

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: sheet,
      valueInputOption: "USER_ENTERED",
      resource: { values },
    });

    await formatSheet(sheet, values);

    res.status(200).json({ message: "Data saved" });
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

function parseContent(data, html, link) {
  const $ = cheerio.load(html);
  const row = [];
  row.push([$(`title`).text().trim(), link, ""]);
  row.push(["", "", ""]);

  $(`${data} div.container div.tickets`).each((_, section) => {
    const $section = $(section);
    const h3Text = $section.find("h3").first().text().trim();
    row.push([h3Text, "date", "link"]);

    // Only get h4, p, and a inside .find-ticket-items
    $section.find(".find-ticket-items").each((_, item) => {
      const h4 = $(item).find("h4").first().text().trim();
      const p = $(item).find("p").first().text().trim();
      const rawHref = $(item).find("a").first().attr("href");
      const href = rawHref ? `=HYPERLINK("${rawHref}", "Click here")` : "";

      row.push([h4, p, href]);
    });

    row.push(["", "", ""]);
  });

  return row;
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
      fields: "userEnteredFormat(horizontalAlignment,wrapStrategy,verticalAlignment)",
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
