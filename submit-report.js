// /api/submit-report.js
// Vercel serverless function: receives the onboarding-test report from
// bea_teacher_onboarding.html and appends a row to a Google Sheet.
//
// SETUP (one-time, ~10 min):
// 1. Create a Google Sheet (e.g. "BEA — Онбординг тичерів") with a
//    tab named "Results". First row (headers) will be created automatically.
// 2. In Google Cloud Console, create a Service Account, enable the
//    "Google Sheets API" for the project, and create a JSON key.
// 3. Share the Google Sheet with the service account's email
//    (found in the JSON key as "client_email") — give it Editor access.
// 4. In your Vercel project → Settings → Environment Variables, add:
//      GOOGLE_SERVICE_ACCOUNT_EMAIL = <client_email from the JSON key>
//      GOOGLE_PRIVATE_KEY           = <private_key from the JSON key>
//                                      (keep the \n escape sequences as-is)
//      SHEET_ID                     = <the long ID from the Sheet's URL>
// 5. Run:  npm install googleapis
// 6. Deploy. The HTML posts to "/api/submit-report" — same origin,
//    no CORS config needed.

import { google } from "googleapis";

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

    // Ensure header row exists (cheap check: read row 1).
    const headerCheck = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Results!A1:G1"
    });
    if (!headerCheck.data.values || headerCheck.data.values.length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: "Results!A1:G1",
        valueInputOption: "RAW",
        requestBody: {
          values: [[
            "Дата", "Ім'я", "Email", "Бал", "Максимум", "Відсоток", "Деталі (JSON)"
          ]]
        }
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Results!A:G",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          submittedAt,
          name,
          email,
          score,
          total,
          pct,
          JSON.stringify({ managerNames, details })
        ]]
      }
    });

    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("submit-report error:", err);
    res.status(500).json({ ok: false, error: err.message });
  }
}
