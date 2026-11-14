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
  if (req.method !== 'POST') {
    return res.status(200).send('Method not allowed');
  }

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

  // Keyboard နေရာတွင် အမြဲပေါ်နေမည့် ခလုတ်များ (Persistent Reply Keyboard)
  const persistentKeyboard = {
    keyboard: [
      [{ text: '✏️ Sheet အသစ်ပြောင်းချိတ်ရန်' }, { text: '📄 လက်ရှိ Sheet ကြည့်ရန်' }],
    ],
    resize_keyboard: true,
    is_persistent: true,
  };

  try {
    const message = req.body?.message;
    if (!message) return res.status(200).send('OK');
    const chatId = message.chat.id;
    const text = (message.text || '').trim();

    // ၁။ /start ရိုက်လျှင် Keyboard ခလုတ်များ ထုတ်ပေးခြင်း
    if (text.startsWith('/start')) {
      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text: '✅ အဆင်သင့်ဖြစ်ပါပြီ။ အောက်ခြေရှိ ခလုတ်များကို နှိပ်၍ အသုံးပြုနိုင်ပါသည်-',
        reply_markup: persistentKeyboard,
      });
      return res.status(200).send('OK');
    }

    // ၂။ "လက်ရှိ Sheet ကြည့်ရန်" ခလုတ်နှိပ်လျှင်
    if (text === '📄 လက်ရှိ Sheet ကြည့်ရန်') {
      const currentId = await getActiveSheetId();
      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text: `📊 *လက်ရှိ အသုံးပြုနေသော Google Sheet:*\nID: \`${currentId}\`\n\n🔗 [Google Sheet ဖွင့်ရန်](https://docs.google.com/spreadsheets/d/${currentId})`,
        parse_mode: 'Markdown',
        disable_web_page_preview: true,
        reply_markup: persistentKeyboard,
      });
      return res.status(200).send('OK');
    }

    // ၃။ "Sheet အသစ်ပြောင်းချိတ်ရန်" ခလုတ်နှိပ်လျှင်
    if (text === '✏️ Sheet အသစ်ပြောင်းချိတ်ရန်') {
      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text:
          `📌 **Sheet အသစ် ပြောင်းချိတ်နည်း**\n\n` +
          `၁။ Sheet အသစ်ရဲ့ **Share** ထဲမှာ အောက်ပါ Email ကို **Editor** ထည့်ပေးပါ (နှိပ်ပြီး Copy ကူးနိုင်သည်)-\n` +
          `\`${SERVICE_EMAIL}\`\n\n` +
          `၂။ ပြီးရင် **Sheet Link** (သို့မဟုတ်) **Sheet ID** ကို ဒီ Chat ထဲကို တိုက်ရိုက် ပို့ပေးလိုက်ပါ-`,
        parse_mode: 'Markdown',
        reply_markup: persistentKeyboard,
      });
      return res.status(200).send('OK');
    }

    // ၄။ Sheet Link သို့မဟုတ် ID တိုက်ရိုက်ပို့လာလျှင် (Reply ပြန်စရာမလိုဘဲ တန်းဖမ်းပေးသည်)
    const possibleSheetId = text && !text.startsWith('/') ? extractSheetId(text) : null;
    if (possibleSheetId) {
      try {
        await setupTargetSheet(possibleSheetId);
      } catch (permErr) {
        await sendTelegram(BOT_TOKEN, {
          chat_id: chatId,
          text:
            `⚠️ **ချိတ်ဆက်မှု မအောင်မြင်ပါ**\n` +
            `Sheet အသစ်မှာ အောက်ပါ Email ကို **Editor** အဖြစ် Share မပေးရသေးပါ-\n\n` +
            `\`${SERVICE_EMAIL}\`\n\n` +
            `Share ပေးပြီးမှ Link သို့မဟုတ် ID ကို ပြန်ပို့ပေးပါ။`,
          parse_mode: 'Markdown',
          reply_markup: persistentKeyboard,
        });
        return res.status(200).send('OK');
      }

      await sheets.spreadsheets.values.update({
        spreadsheetId: MASTER_SHEET_ID,
        range: 'Config!A1',
        valueInputOption: 'RAW',
        requestBody: { values: [[possibleSheetId]] },
      });

      await sendTelegram(BOT_TOKEN, {
        chat_id: chatId,
        text:
          `✅ **Sheet အသစ် ပြောင်းချိတ်ပြီးပါပြီ!**\n\n` +
          `• Sheet ID: \`${possibleSheetId}\`\n` +
          `• \`Deposit\` နှင့် \`Withdraw\` Tab များကို အလိုအလျောက် ပြင်ဆင်ပြီးပါပြီ။\n\n` +
          `Excel ဖိုင်များ စတင်ပို့နိုင်ပါပြီ။`,
        parse_mode: 'Markdown',
        reply_markup: persistentKeyboard,
      });
      return res.status(200).send('OK');
    }

    // ၅။ Excel ဖိုင် ပို့လာလျှင် လက်ရှိ Active Sheet ထဲသို့ ထည့်သွင်းခြင်း
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
            text: `✅ အောင်မြင်ပါသည်။ [${targetTab}] Tab ထဲသို့ စာရင်းအသစ် (${newRows.length}) ကြောင်း ထည့်သွင်းပြီးပါပြီ။${skipMsg}`,
            reply_markup: persistentKeyboard,
          });
        } else {
          await sendTelegram(BOT_TOKEN, {
            chat_id: chatId,
            text: `ℹ️ စာရင်းအသစ် မပါဝင်ပါ။ (စာရင်း ${duplicateCount} ကြောင်းလုံး [${targetTab}] Tab ထဲတွင် ရှိပြီးသား ဖြစ်နေပါသည်)`,
            reply_markup: persistentKeyboard,
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
        reply_markup: persistentKeyboard,
      });
    }
  }

  return res.status(200).send('OK');
}
