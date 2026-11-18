import { google } from 'googleapis';
import { parseDepositExcel, LEDGER_HEADERS } from '../parser.js';

const DISCONNECTED_MARKER = 'DISCONNECTED';
const BANK_REMARKS_TAB = 'BankRemarks';
const BANK_REMARKS_HEADERS = ['Bank', 'Acc No', 'Remark', 'Updated At'];
const PENDING_TAB = 'PendingImports';
const PENDING_HEADERS = ['Batch ID', 'Meta', 'Payload'];
const CHUNK_SIZE = 40000;

function extractSheetId(input) {
  const text = String(input || '').trim();
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) return text;
  return null;
}

function makeBankKey(bank, accNo) {
  return `${String(bank || '').trim().toLowerCase()}|${String(accNo || '').trim().toLowerCase()}`;
}

function getBaseUrl(req) {
  if (process.env.APP_URL) {
    return process.env.APP_URL.replace(/\/+$/, '');
  }
  const host = req.headers['x-forwarded-host'] || req.headers.host || '';
  const proto = req.headers['x-forwarded-proto'] || (host.includes('localhost') || host.startsWith('127.') ? 'http' : 'https');
  return host ? `${proto}://${host}` : '';
}

async function sendTelegram(botToken, payload) {
  if (!botToken) return;
  await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
}

function parseGoogleCreds(raw) {
  if (!raw || !String(raw).trim()) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export default async function handler(req, res) {
  const BOT_TOKEN = process.env.BOT_TOKEN;
  const MASTER_SHEET_ID = process.env.SHEET_ID || '';
  const GOOGLE_CREDS = parseGoogleCreds(process.env.GOOGLE_CREDS_JSON);
  const SERVICE_EMAIL = GOOGLE_CREDS?.client_email || '';

  let sheets = null;
  if (GOOGLE_CREDS) {
    const auth = new google.auth.GoogleAuth({
      credentials: GOOGLE_CREDS,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    sheets = google.sheets({ version: 'v4', auth });
  }

  async function ensureTab(spreadsheetId, title, headers) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheet = meta.data.sheets.find((s) => s.properties.title === title);
    let sheetId = existingSheet?.properties?.sheetId;

    if (!existingSheet) {
      const addRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title } } }] },
      });
      sheetId = addRes.data.replies?.[0]?.addSheet?.properties?.sheetId;
    }

    if (headers && headers.length > 0) {
      const lastCol = String.fromCharCode(64 + headers.length);
      const checkHeader = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${title}!A1:${lastCol}1`,
      });
      const existingRow = checkHeader.data.values?.[0] || [];
      if (existingRow.length === 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${title}!A1:${lastCol}1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [headers] },
        });
      } else if (existingRow.length < headers.length) {
        const merged = headers.map((h, idx) => existingRow[idx] || h);
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${title}!A1:${lastCol}1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [merged] },
        });
      }
    }
    return sheetId;
  }

  async function getActiveSheetId() {
    if (!sheets || !MASTER_SHEET_ID) {
      return '';
    }
    try {
      const cfg = await sheets.spreadsheets.values.get({
        spreadsheetId: MASTER_SHEET_ID,
        range: 'Config!A1',
      });
      const savedId = cfg.data.values?.[0]?.[0]?.trim();
      if (savedId === DISCONNECTED_MARKER) {
        return '';
      }
      return savedId || MASTER_SHEET_ID;
    } catch {
      return MASTER_SHEET_ID;
    }
  }

  async function saveConfigValue(val) {
    await ensureTab(MASTER_SHEET_ID, 'Config', null);
    await sheets.spreadsheets.values.update({
      spreadsheetId: MASTER_SHEET_ID,
      range: 'Config!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [[val]] },
    });
  }

  async function setupTargetSheet(targetSheetId) {
    await ensureTab(targetSheetId, 'Deposit', LEDGER_HEADERS);
    await ensureTab(targetSheetId, 'Withdraw', LEDGER_HEADERS);
  }

  async function getBankRemarks() {
    if (!sheets || !MASTER_SHEET_ID) {
      return { list: [], map: new Map() };
    }
    await ensureTab(MASTER_SHEET_ID, BANK_REMARKS_TAB, BANK_REMARKS_HEADERS);
    const resData = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SHEET_ID,
      range: `${BANK_REMARKS_TAB}!A2:D`,
    });
    const rows = resData.data.values || [];
    const list = [];
    const map = new Map();

    rows.forEach((row, idx) => {
      const bank = String(row[0] || '').trim();
      const accNo = String(row[1] || '').trim();
      const remark = String(row[2] || '').trim();
      const updatedAt = String(row[3] || '').trim();
      if (!bank && !accNo) return;
      const key = makeBankKey(bank, accNo);
      map.set(key, remark);
      list.push({ rowIndex: idx + 2, bank, accNo, remark, updatedAt, key });
    });

    return { list, map };
  }

  async function upsertBankRemarks(entries) {
    await ensureTab(MASTER_SHEET_ID, BANK_REMARKS_TAB, BANK_REMARKS_HEADERS);
    const { list } = await getBankRemarks();
    const existingByKey = new Map(list.map((item) => [item.key, item]));
    const now = new Date().toISOString();
    const toAppend = [];

    for (const entry of entries) {
      const bank = String(entry.bank || '').trim();
      const accNo = String(entry.accNo || '').trim();
      const remark = String(entry.remark || '').trim();
      if (!bank || !accNo || !remark) continue;

      const key = makeBankKey(bank, accNo);
      const existing = existingByKey.get(key);
      if (existing) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: MASTER_SHEET_ID,
          range: `${BANK_REMARKS_TAB}!A${existing.rowIndex}:D${existing.rowIndex}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[bank, accNo, remark, now]] },
        });
        existing.remark = remark;
        existing.updatedAt = now;
      } else {
        toAppend.push([bank, accNo, remark, now]);
        existingByKey.set(key, { bank, accNo, remark, updatedAt: now, key });
      }
    }

    if (toAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: MASTER_SHEET_ID,
        range: `${BANK_REMARKS_TAB}!A:D`,
        valueInputOption: 'RAW',
        requestBody: { values: toAppend },
      });
    }

    return (await getBankRemarks()).list;
  }

  async function deleteBankRemark(bank, accNo) {
    const sheetId = await ensureTab(MASTER_SHEET_ID, BANK_REMARKS_TAB, BANK_REMARKS_HEADERS);
    const { list } = await getBankRemarks();
    const targetKey = makeBankKey(bank, accNo);
    const matches = list.filter((item) => item.key === targetKey).sort((a, b) => b.rowIndex - a.rowIndex);

    if (matches.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: MASTER_SHEET_ID,
        requestBody: {
          requests: matches.map((m) => ({
            deleteDimension: {
              range: {
                sheetId,
                dimension: 'ROWS',
                startIndex: m.rowIndex - 1,
                endIndex: m.rowIndex,
              },
            },
          })),
        },
      });
    }
    return (await getBankRemarks()).list;
  }

  async function savePendingBatch({ batchId, chatId, targetTab, activeSheetId, duplicateCount, rows }) {
    await ensureTab(MASTER_SHEET_ID, PENDING_TAB, PENDING_HEADERS);
    const meta = JSON.stringify({
      chatId,
      targetTab,
      activeSheetId,
      duplicateCount,
      createdAt: new Date().toISOString(),
    });
    const serialized = JSON.stringify(rows);
    const chunks = [];
    for (let i = 0; i < serialized.length; i += CHUNK_SIZE) {
      chunks.push(serialized.slice(i, i + CHUNK_SIZE));
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: MASTER_SHEET_ID,
      range: `${PENDING_TAB}!A:Z`,
      valueInputOption: 'RAW',
      requestBody: { values: [[batchId, meta, ...chunks]] },
    });
  }

  async function getPendingBatch(batchId) {
    if (!sheets || !MASTER_SHEET_ID) return null;
    await ensureTab(MASTER_SHEET_ID, PENDING_TAB, PENDING_HEADERS);
    const resData = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SHEET_ID,
      range: `${PENDING_TAB}!A2:Z`,
    });
    const rows = resData.data.values || [];
    if (rows.length === 0) return null;

    let targetIndex = -1;
    if (batchId) {
      targetIndex = rows.findIndex((r) => String(r[0] || '').trim() === batchId);
    } else {
      targetIndex = rows.length - 1;
    }
    if (targetIndex === -1) return null;

    const row = rows[targetIndex];
    const id = String(row[0] || '').trim();
    const meta = JSON.parse(row[1] || '{}');
    const serialized = row.slice(2).join('');
    const parsedRows = JSON.parse(serialized || '[]');

    return {
      rowIndex: targetIndex + 2,
      batchId: id,
      chatId: meta.chatId,
      targetTab: meta.targetTab || 'Deposit',
      activeSheetId: meta.activeSheetId,
      duplicateCount: meta.duplicateCount || 0,
      createdAt: meta.createdAt,
      rows: parsedRows,
    };
  }

  async function deletePendingBatch(batchId) {
    const sheetId = await ensureTab(MASTER_SHEET_ID, PENDING_TAB, PENDING_HEADERS);
    const resData = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SHEET_ID,
      range: `${PENDING_TAB}!A2:A`,
    });
    const rows = resData.data.values || [];
    const indices = [];
    rows.forEach((r, idx) => {
      if (String(r[0] || '').trim() === batchId) {
        indices.push(idx + 2);
      }
    });
    if (indices.length === 0) return;

    indices.sort((a, b) => b - a);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: MASTER_SHEET_ID,
      requestBody: {
        requests: indices.map((rowIndex) => ({
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex,
            },
          },
        })),
      },
    });
  }

  async function deduplicateRows(spreadsheetId, targetTab, candidateRows) {
    await ensureTab(spreadsheetId, targetTab, LEDGER_HEADERS);
    const rangeName = `${targetTab}!A:H`;
    const existingRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
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

    for (const row of candidateRows) {
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

    return { newRows, duplicateCount };
  }

  function summarizeAccounts(rows, remarksMap) {
    const grouped = new Map();
    for (const row of rows) {
      const bank = String(row[1] || 'Auto').trim();
      const accNo = String(row[2] || '0000').trim();
      const key = makeBankKey(bank, accNo);
      const savedRemark = remarksMap.get(key) || '';
      if (!grouped.has(key)) {
        grouped.set(key, {
          bank,
          accNo,
          key,
          count: 0,
          remark: savedRemark,
          isNew: !savedRemark,
        });
      }
      grouped.get(key).count += 1;
    }
    return Array.from(grouped.values());
  }

  // ၁။ Mini App မှ Config / Remarks / Pending Batch များ တောင်းယူခြင်း (GET)
  if (req.method === 'GET') {
    const action = req.query?.action;

    if (action === 'config') {
      const activeSheetId = await getActiveSheetId();
      const latestPending = await getPendingBatch();
      return res.status(200).json({
        ok: true,
        connected: Boolean(activeSheetId),
        envConfigured: Boolean(sheets && MASTER_SHEET_ID),
        activeSheetId: activeSheetId || '',
        serviceEmail: SERVICE_EMAIL || 'Not configured (set GOOGLE_CREDS_JSON)',
        latestPendingBatchId: latestPending?.batchId || '',
      });
    }

    if (action === 'remarks') {
      try {
        const { list } = await getBankRemarks();
        return res.status(200).json({ ok: true, remarks: list });
      } catch (err) {
        return res.status(500).json({ ok: false, error: err.message });
      }
    }

    if (action === 'pending_batch') {
      try {
        const batch = await getPendingBatch(req.query?.batchId);
        if (!batch) {
          return res.status(200).json({ ok: true, batch: null });
        }
        const { map } = await getBankRemarks();
        const accounts = summarizeAccounts(batch.rows, map);
        return res.status(200).json({
          ok: true,
          batch: {
            batchId: batch.batchId,
            targetTab: batch.targetTab,
            totalRows: batch.rows.length,
            duplicateCount: batch.duplicateCount,
            createdAt: batch.createdAt,
            accounts,
          },
        });
      } catch (err) {
        return res.status(500).json({ ok: false, error: err.message });
      }
    }

    return res.status(200).send('Bot Webhook Active');
  }

  if (req.method !== 'POST') {
    return res.status(200).send('Method not allowed');
  }

  // ၂။ Mini App မှ Sheet ID အသစ် ပြောင်းလဲသိမ်းဆည်းခြင်း
  if (req.body?.action === 'update_sheet') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({
        ok: false,
        error: 'Google Credentials (GOOGLE_CREDS_JSON) နှင့် SHEET_ID ကို Environment Variables တွင် အရင်ထည့်သွင်းပါ။',
      });
    }
    const newId = extractSheetId(req.body.sheetInput);
    if (!newId) {
      return res.status(400).json({ ok: false, error: 'Sheet Link သို့မဟုတ် ID ပုံစံ မှားယွင်းနေပါသည်။' });
    }
    try {
      await setupTargetSheet(newId);
      await saveConfigValue(newId);
      return res.status(200).json({ ok: true, connected: true, activeSheetId: newId });
    } catch (err) {
      return res.status(403).json({
        ok: false,
        error: 'ထို Sheet ကို ဝင်ခွင့်မရှိပါ။ Service Email ကို Editor အဖြစ် Share အရင်ပေးပါ။',
      });
    }
  }

  // ၂.၅။ လက်ရှိ ချိတ်ဆက်ထားသော Sheet ကို Disconnect လုပ်ခြင်း
  if (req.body?.action === 'disconnect_sheet') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({
        ok: false,
        error: 'Google Credentials (GOOGLE_CREDS_JSON) နှင့် SHEET_ID ကို Environment Variables တွင် အရင်ထည့်သွင်းပါ။',
      });
    }
    try {
      await saveConfigValue(DISCONNECTED_MARKER);
      return res.status(200).json({ ok: true, connected: false, activeSheetId: '' });
    } catch (err) {
      return res.status(500).json({
        ok: false,
        error: `Disconnect လုပ်ရာတွင် အမှားအယွင်းရှိသည်: ${err.message}`,
      });
    }
  }

  // ၂.၆။ Saved Bank Remark အသစ်ထည့်/ပြင်ဆင်ခြင်း
  if (req.body?.action === 'upsert_remark') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({
        ok: false,
        error: 'Google Credentials (GOOGLE_CREDS_JSON) နှင့် SHEET_ID ကို Environment Variables တွင် အရင်ထည့်သွင်းပါ။',
      });
    }
    const bank = String(req.body.bank || '').trim();
    const accNo = String(req.body.accNo || '').trim();
    const remark = String(req.body.remark || '').trim();
    if (!bank || !accNo || !remark) {
      return res.status(400).json({ ok: false, error: 'Bank, Acc No နှင့် Remark အားလုံး ဖြည့်သွင်းပေးပါ။' });
    }
    try {
      const remarks = await upsertBankRemarks([{ bank, accNo, remark }]);
      return res.status(200).json({ ok: true, remarks });
    } catch (err) {
      return res.status(500).json({ ok: false, error: err.message });
    }
  }

  // ၂.၇။ Saved Bank Remark ဖျက်ခြင်း
  if (req.body?.action === 'delete_remark') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({
        ok: false,
        error: 'Google Credentials (GOOGLE_CREDS_JSON) နှင့် SHEET_ID ကို Environment Variables တွင် အရင်ထည့်သွင်းပါ။',
      });
    }
    const bank = String(req.body.bank || '').trim();
    const accNo = String(req.body.accNo || '').trim();
    try {
      const remarks = await deleteBankRemark(bank, accNo);
      return res.status(200).json({ ok: true, remarks });
    } catch (err) {
      return res.status(500).json({ ok: false, error: err.message });
    }
  }

  // ၂.၈။ Pending Excel Batch အတွက် Remark ဖြည့်ပြီး Sheet ထဲသိမ်းခြင်း
  if (req.body?.action === 'submit_batch_remarks') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({
        ok: false,
        error: 'Google Credentials (GOOGLE_CREDS_JSON) နှင့် SHEET_ID ကို Environment Variables တွင် အရင်ထည့်သွင်းပါ။',
      });
    }
    const { batchId, remarks: submittedRemarks = [] } = req.body;
    try {
      const batch = await getPendingBatch(batchId);
      if (!batch) {
        return res.status(404).json({ ok: false, error: 'Pending စာရင်း ရှာမတွေ့ပါ (သို့မဟုတ် သိမ်းပြီးသား ဖြစ်နေပါသည်)။' });
      }

      const targetSheetId = batch.activeSheetId || (await getActiveSheetId());
      if (!targetSheetId) {
        return res.status(400).json({ ok: false, error: 'Google Sheet ချိတ်ဆက်ထားခြင်း မရှိပါ။' });
      }

      const toSave = submittedRemarks.filter((item) => item.saveForFuture !== false && String(item.remark || '').trim());
      const updatedRemarksList = toSave.length > 0 ? await upsertBankRemarks(toSave) : (await getBankRemarks()).list;

      const remarkMap = new Map(updatedRemarksList.map((r) => [r.key, r.remark]));
      for (const item of submittedRemarks) {
        const key = makeBankKey(item.bank, item.accNo);
        const text = String(item.remark || '').trim();
        if (text) remarkMap.set(key, text);
      }

      const { newRows, duplicateCount: extraDup } = await deduplicateRows(targetSheetId, batch.targetTab, batch.rows);
      const rowsWithRemarks = newRows.map((row) => {
        const key = makeBankKey(row[1], row[2]);
        return [...row.slice(0, 7), remarkMap.get(key) || ''];
      });

      if (rowsWithRemarks.length > 0) {
        await sheets.spreadsheets.values.append({
          spreadsheetId: targetSheetId,
          range: `${batch.targetTab}!A:H`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: rowsWithRemarks },
        });
      }

      await deletePendingBatch(batch.batchId);

      const totalSkipped = (batch.duplicateCount || 0) + extraDup;
      if (batch.chatId) {
        const skipMsg = totalSkipped > 0 ? `\n(စာရင်းဟောင်း ${totalSkipped} ကြောင်း ထပ်နေ၍ ကျော်လိုက်ပါသည်)` : '';
        await sendTelegram(BOT_TOKEN, {
          chat_id: batch.chatId,
          text: rowsWithRemarks.length > 0
            ? `✅ [${batch.targetTab}] ထဲသို့ Remark နှင့်အတူ စာရင်းအသစ် (${rowsWithRemarks.length}) ကြောင်း ထည့်ပြီးပါပြီ။${skipMsg}`
            : `ℹ️ စာရင်းအသစ် မပါဝင်ပါ။ (${totalSkipped} ကြောင်းလုံး [${batch.targetTab}] ထဲတွင် ရှိပြီးသားဖြစ်သည်)`,
        });
      }

      return res.status(200).json({
        ok: true,
        insertedCount: rowsWithRemarks.length,
        totalSkipped,
        targetTab: batch.targetTab,
        remarks: updatedRemarksList,
      });
    } catch (err) {
      return res.status(500).json({ ok: false, error: err.message });
    }
  }

  // ၂.၉။ Pending Excel Batch ကို ပယ်ဖျက်ခြင်း
  if (req.body?.action === 'discard_batch') {
    if (!sheets || !MASTER_SHEET_ID) {
      return res.status(400).json({ ok: false, error: 'Google Credentials မပြည့်စုံသေးပါ။' });
    }
    try {
      const batch = await getPendingBatch(req.body.batchId);
      if (batch) {
        await deletePendingBatch(batch.batchId);
        if (batch.chatId) {
          await sendTelegram(BOT_TOKEN, {
            chat_id: batch.chatId,
            text: `🗑️ [${batch.targetTab}] Excel စာရင်းသွင်းခြင်းကို ပယ်ဖျက်လိုက်ပါသည်။`,
          });
        }
      }
      return res.status(200).json({ ok: true });
    } catch (err) {
      return res.status(500).json({ ok: false, error: err.message });
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
      if (!activeSheetId) {
        await sendTelegram(BOT_TOKEN, {
          chat_id: chatId,
          text: '⚠️ Google Sheet ချိတ်ဆက်ထားခြင်း မရှိပါ။ Sheet Manager မှတစ်ဆင့် Sheet အသစ် ချိတ်ဆက်ပေးပါ။',
        });
        return res.status(200).send('OK');
      }

      const fileInfoRes = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/getFile?file_id=${message.document.file_id}`);
      const fileInfo = await fileInfoRes.json();
      const fileRes = await fetch(`https://api.telegram.org/file/bot${BOT_TOKEN}/${fileInfo.result.file_path}`);
      const arrayBuffer = await fileRes.arrayBuffer();

      const parsedRows = parseDepositExcel(arrayBuffer);

      if (parsedRows.length > 0) {
        const isWithdraw = parsedRows[0][3] === 'Withdraw';
        const targetTab = isWithdraw ? 'Withdraw' : 'Deposit';

        const { newRows, duplicateCount } = await deduplicateRows(activeSheetId, targetTab, parsedRows);

        if (newRows.length === 0) {
          await sendTelegram(BOT_TOKEN, {
            chat_id: chatId,
            text: `ℹ️ စာရင်းအသစ် မပါဝင်ပါ။ (${duplicateCount} ကြောင်းလုံး [${targetTab}] ထဲတွင် ရှိပြီးသားဖြစ်သည်)`,
          });
          return res.status(200).send('OK');
        }

        const { map: remarksMap } = await getBankRemarks();
        const accountSummary = summarizeAccounts(newRows, remarksMap);
        const missingAccounts = accountSummary.filter((acc) => acc.isNew);

        if (missingAccounts.length === 0) {
          const rowsWithRemarks = newRows.map((row) => {
            const key = makeBankKey(row[1], row[2]);
            return [...row.slice(0, 7), remarksMap.get(key) || ''];
          });

          await sheets.spreadsheets.values.append({
            spreadsheetId: activeSheetId,
            range: `${targetTab}!A:H`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: rowsWithRemarks },
          });

          const skipMsg = duplicateCount > 0 ? `\n(စာရင်းဟောင်း ${duplicateCount} ကြောင်း ထပ်နေ၍ ကျော်လိုက်ပါသည်)` : '';
          await sendTelegram(BOT_TOKEN, {
            chat_id: chatId,
            text: `✅ [${targetTab}] ထဲသို့ Auto-Remark ဖြင့် စာရင်းအသစ် (${rowsWithRemarks.length}) ကြောင်း ထည့်ပြီးပါပြီ။${skipMsg}`,
          });
        } else {
          const batchId = `batch_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
          await savePendingBatch({
            batchId,
            chatId,
            targetTab,
            activeSheetId,
            duplicateCount,
            rows: newRows,
          });

          const accLines = missingAccounts
            .map((a) => `• ${a.bank} - ${a.accNo} (${a.count} ကြောင်း)`)
            .join('\n');

          const baseUrl = getBaseUrl(req);
          const miniAppUrl = baseUrl ? `${baseUrl}/?batchId=${encodeURIComponent(batchId)}` : '';
          const payload = {
            chat_id: chatId,
            text: `📝 [${targetTab}] အတွက် Remark မရှိသေးသော Bank Acc (${missingAccounts.length}) ခု တွေ့ရှိပါသည်:\n${accLines}\n\nအောက်ပါ Mini App ခလုတ်ကို နှိပ်၍ Remark ထည့်သွင်းသိမ်းဆည်းပါ။`,
          };
          if (miniAppUrl.startsWith('https://')) {
            payload.reply_markup = {
              inline_keyboard: [[{ text: '✏️ Mini App ဖွင့်၍ Remark ထည့်ရန်', web_app: { url: miniAppUrl } }]],
            };
          }
          await sendTelegram(BOT_TOKEN, {
            ...payload,
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
