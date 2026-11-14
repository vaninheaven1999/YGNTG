import { google } from 'googleapis';
import { parseDepositExcel, LEDGER_HEADERS } from '../parser.js';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(200).send('Method not allowed');
  }

  const BOT_TOKEN = process.env.BOT_TOKEN;
  const SHEET_ID = process.env.SHEET_ID;
  const GOOGLE_CREDS = JSON.parse(process.env.GOOGLE_CREDS_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  const message = req.body?.message;
  if (message && message.document) {
    const chatId = message.chat.id;
    const fileName = message.document.file_name || '';

    try {
      // Telegram မှ File Download ဆွဲခြင်း
      const fileInfoRes = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/getFile?file_id=${message.document.file_id}`);
      const fileInfo = await fileInfoRes.json();
      const fileRes = await fetch(`https://api.telegram.org/file/bot${BOT_TOKEN}/${fileInfo.result.file_path}`);
      const arrayBuffer = await fileRes.arrayBuffer();

      // Data ခွဲထုတ်ခြင်း
      const sheetValues = parseDepositExcel(arrayBuffer);

      if (sheetValues.length > 0) {
        // Sheet ထဲ ထည့်ခြင်း
        await sheets.spreadsheets.values.append({
          spreadsheetId: SHEET_ID,
          range: 'Sheet1!A:G',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: sheetValues },
        });

        // Telegram သို့ အကြောင်းပြန်ခြင်း
        await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            chat_id: chatId,
            text: `✅ အောင်မြင်ပါသည်။ စာရင်း (${sheetValues.length}) ကြောင်းကို Google Sheet ထဲ ထည့်သွင်းပြီးပါပြီ။`
          }),
        });
      }
    } catch (err) {
      await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ chat_id: chatId, text: `⚠️ အမှားအယွင်းရှိသည်: ${err.message}` }),
      });
    }
  }

  return res.status(200).send('OK');
}
