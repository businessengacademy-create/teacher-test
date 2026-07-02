// /api/submit-report.js
// Vercel serverless function: receives the onboarding-test report from
// index.html and writes it to a Google Sheet in two places:
//   - "Results" tab: one summary row per submission (score, %, etc.)
//   - "Answers" tab: one row per question per submission, so you can
//     see exactly which questions each teacher got right/wrong, and
//     filter/sort to spot commonly-missed questions across everyone.
//
// SETUP (one-time):
// 1. Google Sheet with a tab named "Results" (the "Answers" tab is
//    created automatically on first submission if it doesn't exist).
// 2. Google Cloud: Service Account + Sheets API enabled + JSON key.
// 3. Share the Sheet with the service account's email — Editor access.
// 4. Vercel env vars: GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY,
//    SHEET_ID.
// 5. package.json must list "googleapis" as a dependency.

import { google } from "googleapis";

const RESULTS_HEADER = ["Дата", "Ім'я", "Email", "Бал", "Максимум", "Відсоток", "Менеджери проєктів"];
const ANSWERS_HEADER = ["Дата", "Ім'я", "Email", "№", "Розділ", "Питання", "Тип", "Відповідь тичера", "Правильна відповідь", "Правильно?"];

async function ensureSheetExists(sheets, spreadsheetId, title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const exists = (meta.data.sheets || []).some(s => s.properties.title === title);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] }
    });
  }
}

async function ensureHeader(sheets, spreadsheetId, title, header) {
  const range = `${title}!A1:${String.fromCharCode(64 + header.length)}1`;
  const check = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  if (!check.data.values || check.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: "RAW",
      requestBody: { values: [header] }
    });
  }
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const body = req.body || {};
    const {
      name = "",
      email = "",
      submittedAt = new Date().toISOString(),
      score = 0,
      total = 0,
      pct = 0,
      managerNames = "",
      details = []
    } = body;

    const auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );

    const sheets = google.sheets({ version: "v4", auth });
    const spreadsheetId = process.env.SHEET_ID;

    // --- Results (summary row) ---
    await ensureSheetExists(sheets, spreadsheetId, "Results");
    await ensureHeader(sheets, spreadsheetId, "Results", RESULTS_HEADER);
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Results!A:G",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[submittedAt, name, email, score, total, pct, managerNames]]
      }
    });

    // --- Answers (one row per question) ---
    if (Array.isArray(details) && details.length > 0) {
      await ensureSheetExists(sheets, spreadsheetId, "Answers");
      await ensureHeader(sheets, spreadsheetId, "Answers", ANSWERS_HEADER);

      const rows = details.map((d, i) => [
        submittedAt,
        name,
        email,
        i + 1,
        d.section || "",
        d.question || "",
        d.type || "",
        d.userAnswer || "",
        d.correctAnswer || "",
        d.isCorrect === null ? "—" : (d.isCorrect ? "✅" : "❌")
      ]);

      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "Answers!A:J",
        valueInputOption: "RAW",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: rows }
      });
    }

    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("submit-report error:", err);
    res.status(500).json({ ok: false, error: err.message });
  }
}
