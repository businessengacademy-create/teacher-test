// /api/submit-feedback.js
// Vercel serverless function: receives anonymous onboarding feedback
// (star rating + free text) and appends it to a "Feedback" tab in the
// same Google Sheet. Deliberately does NOT receive or store any name
// or email — keep it that way to preserve anonymity.

import { google } from "googleapis";

const FEEDBACK_HEADER = ["Дата", "Оцінка (1-5)", "Фідбек"];

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
    const rating = Number(body.rating) || 0;
    const feedback = (body.feedback || "").toString().slice(0, 5000);
    const submittedAt = body.submittedAt || new Date().toISOString();

    if (!rating && !feedback) {
      res.status(400).json({ ok: false, error: "Empty feedback" });
      return;
    }

    const auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );

    const sheets = google.sheets({ version: "v4", auth });
    const spreadsheetId = process.env.SHEET_ID;

    await ensureSheetExists(sheets, spreadsheetId, "Feedback");
    await ensureHeader(sheets, spreadsheetId, "Feedback", FEEDBACK_HEADER);

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Feedback!A:C",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[submittedAt, rating || "", feedback]]
      }
    });

    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("submit-feedback error:", err);
    res.status(500).json({ ok: false, error: err.message });
  }
}
