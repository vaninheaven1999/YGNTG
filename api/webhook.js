import { google } from 'googleapis';
import { parseDepositExcel, LEDGER_HEADERS } from '../parser.js';

function extractSheetId(input) {
  const text = String(input || '').trim();
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) return text;
  return null;
}

async function sendTelegram(botToken, payload) {
  await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
}

export default async function handler(req, res) {
  const BOT_TOKEN = process.env.BOT_TOKEN;
  const MASTER_SHEET_ID = process.env.SHEET_ID;
  const GOOGLE_CREDS = JSON.parse(process.env.GOOGLE_CREDS_JSON);
  const SERVICE_EMAIL = GOOGLE_CREDS.client_email;

  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  async function getActiveSheetId() {
    try {
      const cfg = await sheets.spreadsheets.values.get({
        spreadsheetId: MASTER_SHEET_ID,
        range: 'Config!A1',
      });
      const savedId = cfg.data.values?.[0]?.[0]?.trim();
      return savedId || MASTER_SHEET_ID;
    } catch {
      return MASTER_SHEET_ID;
    }
  }

  async function setupTargetSheet(targetSheetId) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: targetSheetId });
    const existingTabs = meta.data.sheets.map((s) => s.properties.title);
    const requiredTabs = ['Deposit', 'Withdraw'];
    const requests = [];

    for (const tab of requiredTabs) {
      if (!existingTabs.includes(tab)) {
        requests.push({ addSheet: { properties: { title: tab } } });
      }
    }

    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: targetSheetId,
        requestBody: { requests },
      });
    }

    for (const tab of requiredTabs) {
      const checkHeader = await sheets.spreadsheets.values.get({
        spreadsheetId: targetSheetId,
        range: `${tab}!A1:G1`,
      });
      if (!checkHeader.data.values || checkHeader.data.values.length === 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: targetSheetId,
          range: `${tab}!A1:G1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [LEDGER_HEADERS] },
        });
      }
    }
  }

  // ၁။ Mini App မှ လက်ရှိ Config အချက်အလက်များ တောင်းယူခြင်း (GET)
  if (req.method === 'GET') {
    if (req.query?.action === 'config') {
      const activeSheetId = await getActiveSheetId();
      return res.status(200).json({ ok: true, activeSheetId, serviceEmail: SERVICE_EMAIL });
    }
    return res.status(200).send('Bot Webhook Active');
  }

  if (req.method !== 'POST') {
    return res.status(200).send('Method not allowed');
  }

  // ၂။ Mini App မှ Sheet ID အသစ် ပြောင်းလဲသိမ်းဆည်းခြင်း
  if (req.body?.action === 'update_sheet') {
    const newId = extractSheetId(req.body.sheetInput);
    if (!newId) {
      return res.status(400).json({ ok: false, error: 'Sheet Link သို့မဟုတ် ID ပုံစံ မှားယွင်းနေပါသည်။' });
    }
    try {
      await setupTargetSheet(newId);
      await sheets.spreadsheets.values.update({
        spreadsheetId: MASTER_SHEET_ID,
        range: 'Config!A1',
        valueInputOption: 'RAW',
        requestBody: { values: [[newId]] },
      });
      return res.status(200).json({ ok: true, activeSheetId: newId });
    } catch (err) {
      return res.status(403).json({
        ok: false,
        error: 'ထို Sheet ကို ဝင်ခွင့်မရှိပါ။ Service Email ကို Editor အဖြစ် Share အရင်ပေးပါ။',
      });
    }
  }

  // ၃။ Telegram Bot မှ Message များ လက်ခံဆောင်ရွက်ခြင်း
  try {
    const message = req.body?.message;
    if (!message) return res.status(200).send('OK');
    const chatId = message.chat.id;
    const text = (message.text || '').trim();

    if (text.startsWith('/start')) {
      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text: '✅ အဆင်သင့်ဖြစ်ပါပြီ။ Excel ဖိုင်များကို တိုက်ရိုက် ပို့ပေးနိုင်ပါပြီ။',
        reply_markup: { remove_keyboard: true },
      });
      return res.status(200).send('OK');
    }

    if (message.document) {
      const activeSheetId = await getActiveSheetId();

      const fileInfoRes = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/getFile?file_id=${message.document.file_id}`);
      const fileInfo = await fileInfoRes.json();
      const fileRes = await fetch(`https://api.telegram.org/file/bot${BOT_TOKEN}/${fileInfo.result.file_path}`);
      const arrayBuffer = await fileRes.arrayBuffer();

      const parsedRows = parseDepositExcel(arrayBuffer);

      if (parsedRows.length > 0) {
        const isWithdraw = parsedRows[0][3] === 'Withdraw';
        const targetTab = isWithdraw ? 'Withdraw' : 'Deposit';
        const rangeName = `${targetTab}!A:G`;

        const existingRes = await sheets.spreadsheets.values.get({
          spreadsheetId: activeSheetId,
          range: rangeName,
        });
        const existingRows = existingRes.data.values || [];

        const existingKeys = new Set(
          existingRows.map((row) => {
            const date = String(row[0] || '').trim();
            const type = String(row[3] || '').trim();
            const username = String(row[4] || '').trim().toLowerCase();
            const time = String(row[5] || '').trim();
            const amount = Number(String(row[6] || '0').replace(/,/g, '')).toFixed(2);
            return `${date}|${type}|${username}|${time}|${amount}`;
          })
        );

        const newRows = [];
        let duplicateCount = 0;

        for (const row of parsedRows) {
          const date = String(row[0] || '').trim();
          const type = String(row[3] || '').trim();
          const username = String(row[4] || '').trim().toLowerCase();
          const time = String(row[5] || '').trim();
          const amount = Number(row[6] || 0).toFixed(2);
          const key = `${date}|${type}|${username}|${time}|${amount}`;

          if (existingKeys.has(key)) {
            duplicateCount++;
          } else {
            existingKeys.add(key);
            newRows.push(row);
          }
        }

        if (newRows.length > 0) {
          await sheets.spreadsheets.values.append({
            spreadsheetId: activeSheetId,
            range: rangeName,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: newRows },
          });

          const skipMsg = duplicateCount > 0 ? `\n(စာရင်းဟောင်း ${duplicateCount} ကြောင်း ထပ်နေ၍ ကျော်လိုက်ပါသည်)` : '';
          await sendTelegram(BOT_TOKEN, {
            chat_id: chatId,
            text: `✅ [${targetTab}] ထဲသို့ စာရင်းအသစ် (${newRows.length}) ကြောင်း ထည့်ပြီးပါပြီ။${skipMsg}`,
          });
        } else {
          await sendTelegram(BOT_TOKEN, {
            chat_id: chatId,
            text: `ℹ️ စာရင်းအသစ် မပါဝင်ပါ။ (${duplicateCount} ကြောင်းလုံး [${targetTab}] ထဲတွင် ရှိပြီးသားဖြစ်သည်)`,
          });
        }
      }
    }
  } catch (err) {
    const chatId = req.body?.message?.chat?.id;
    if (chatId) {
      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text: `⚠️ အမှားအယွင်းရှိသည်: ${err.message}`,
      });
    }
  }

  return res.status(200).send('OK');
}
