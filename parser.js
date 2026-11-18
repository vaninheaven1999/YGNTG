import * as XLSX from 'xlsx';

export const LEDGER_HEADERS = ['Date', 'Bank', 'Acc No', 'Type', 'Username', 'Time', 'Amount'];

function parseNumeric(val) {
  if (typeof val === 'number') return Number.isNaN(val) ? 0 : val;
  if (!val) return 0;
  const num = parseFloat(String(val).replace(/,/g, '').trim());
  return Number.isNaN(num) ? 0 : num;
}

function formatAccountNo(val) {
  if (val == null) return '';
  if (typeof val === 'number') {
    return Number.isInteger(val) ? val.toLocaleString('fullwide', { useGrouping: false }) : String(val).trim();
  }
  return String(val).trim();
}

function formatDateOnly(val) {
  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    const d = String(val.getDate()).padStart(2, '0');
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const y = val.getFullYear();
    return `${d}/${m}/${y}`;
  }
  const str = String(val ?? '').trim();
  return str.split(/\s+/)[0] || '';
}

function formatTimeOnly(val) {
  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    const hr = String(val.getHours()).padStart(2, '0');
    const min = String(val.getMinutes()).padStart(2, '0');
    const sec = String(val.getSeconds()).padStart(2, '0');
    return `${hr}:${min}:${sec}`;
  }
  const str = String(val ?? '').trim();
  const parts = str.split(/\s+/);
  const candidate = parts.length > 1 ? parts[1] : parts[0];
  return candidate && candidate.includes(':') ? candidate : '';
}

function formatExcelDateTime(val) {
  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    const d = String(val.getDate()).padStart(2, '0');
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const y = val.getFullYear();
    const hr = String(val.getHours()).padStart(2, '0');
    const min = String(val.getMinutes()).padStart(2, '0');
    const sec = String(val.getSeconds()).padStart(2, '0');
    return `${d}/${m}/${y} ${hr}:${min}:${sec}`;
  }
  return String(val ?? '').trim();
}

export function extractDateAndTime(dtStr) {
  if (!dtStr) return { dateStr: '10-1-2026', timeStr: '0:00:00', timeSeconds: 0 };
  const parts = String(dtStr).trim().split(/\s+/);
  const datePart = parts[0] || '';
  const timePart = parts[1] || '00:00:00';
  let day = 1;
  let month = 10;
  let year = 2026;

  if (datePart.includes('/') || datePart.includes('-')) {
    const sep = datePart.includes('/') ? '/' : '-';
    const dp = datePart.split(sep).map((s) => parseInt(s, 10));
    if (dp.length >= 3 && !Number.isNaN(dp[0]) && !Number.isNaN(dp[1]) && !Number.isNaN(dp[2])) {
      if (dp[0] > 31) {
        // YYYY-MM-DD or YYYY/MM/DD
        year = dp[0];
        month = dp[1] || 1;
        day = dp[2] || 1;
      } else {
        // DD/MM/YYYY or DD-MM-YYYY
        day = dp[0] || 1;
        month = dp[1] || 1;
        year = dp[2] < 100 ? 2000 + dp[2] : dp[2] || 2026;
      }
    }
  }

  const formattedDate = `${day}-${month}-${year}`;
  const tp = timePart.split(':');
  const hr = parseInt(tp[0] || '0', 10) || 0;
  const min = parseInt(tp[1] || '0', 10) || 0;
  const sec = parseInt(tp[2] || '0', 10) || 0;
  return {
    dateStr: formattedDate,
    timeStr: `${hr}:${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')}`,
    timeSeconds: hr * 3600 + min * 60 + sec,
  };
}

export function parseSheetDateTimeMs(dateVal, timeVal) {
  const dStr = String(dateVal ?? '').trim();
  const tStr = String(timeVal ?? '').trim();

  let year = 2026;
  let month = 1;
  let day = 1;

  if (dStr.includes('-') || dStr.includes('/')) {
    const sep = dStr.includes('-') ? '-' : '/';
    const parts = dStr.split(sep).map((s) => parseInt(s, 10));
    if (parts.length >= 3 && !Number.isNaN(parts[0]) && !Number.isNaN(parts[1]) && !Number.isNaN(parts[2])) {
      if (parts[0] > 31) {
        year = parts[0];
        month = parts[1] || 1;
        day = parts[2] || 1;
      } else {
        day = parts[0] || 1;
        month = parts[1] || 1;
        year = parts[2] < 100 ? 2000 + parts[2] : parts[2] || 2026;
      }
    }
  }

  let hr = 0;
  let min = 0;
  let sec = 0;
  if (tStr) {
    const upper = tStr.toUpperCase();
    const isPM = upper.includes('PM');
    const isAM = upper.includes('AM');
    const cleanTime = upper.replace(/AM|PM/g, '').trim();
    const tp = cleanTime.split(':').map((s) => parseInt(s, 10));
    hr = tp[0] || 0;
    min = tp[1] || 0;
    sec = tp[2] || 0;
    if (isPM && hr < 12) hr += 12;
    if (isAM && hr === 12) hr = 0;
  }

  return Date.UTC(year, month - 1, day, hr, min, sec);
}

function isHeaderKeywordCell(cell) {
  if (typeof cell !== 'string') return false;
  const raw = cell.trim();
  if (!raw || raw.length > 40) return false;
  const c = raw.toLowerCase();
  return (
    raw === 'จาก' ||
    raw.includes('วันที่') ||
    raw.includes('วันเวลา') ||
    raw === 'เวลา' ||
    c.includes('date') ||
    c.includes('time') ||
    c.includes('username') ||
    c === 'user' ||
    raw.includes('ยูสเซอร์') ||
    raw.includes('ธนาคาร') ||
    c.includes('bank') ||
    raw.includes('เลขบัญชี') ||
    raw.includes('เลขที่บัญชี') ||
    c.includes('acc') ||
    raw.includes('ยอดฝาก') ||
    raw.includes('ยอดถอน') ||
    raw.includes('จำนวนเงิน') ||
    raw.includes('ยอดเงิน') ||
    c.includes('deposit') ||
    c.includes('withdraw') ||
    c.includes('amount')
  );
}

export function parseDepositExcel(buffer) {
  const input = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : buffer;
  const workbook = XLSX.read(input, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', blankrows: false });
    return rows.length > 0;
  }) || workbook.SheetNames[0];
  if (!sheetName) return [];

  const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', blankrows: false });
  if (!jsonData.length) return [];

  let headerIndex = 0;
  let bestScore = 0;
  const maxScan = Math.min(10, jsonData.length);
  for (let i = 0; i < maxScan; i++) {
    const row = jsonData[i] || [];
    const score = row.reduce((acc, cell) => acc + (isHeaderKeywordCell(cell) ? 1 : 0), 0);
    if (score > bestScore) {
      bestScore = score;
      headerIndex = i;
    }
  }

  const headers = jsonData[headerIndex].map((h) => String(h || '').trim());
  const isWithdrawFile = headers.some((h) => h.includes('ยอดถอน') || h.toLowerCase().includes('withdraw'));
  const hasJakColumn = headers.some((h) => h === 'จาก');

  const rawRows = jsonData.slice(headerIndex + 1).map((row) => {
    const obj = {};
    const rowText = row.map((c) => String(c ?? '').trim().toLowerCase()).join(' ');
    if (rowText.startsWith('รวม') || rowText.startsWith('total')) {
      return obj;
    }

    headers.forEach((h, idx) => {
      if (!h) return;
      const val = row[idx];
      const lowerH = h.toLowerCase();

      if (h.includes('วันที่') || h.includes('วันเวลา') || h.includes('วัน/เวลา') || lowerH.includes('date')) {
        const dt = formatExcelDateTime(val);
        obj.datePart = formatDateOnly(val);
        const maybeTime = formatTimeOnly(val);
        if (maybeTime && !obj.timePart) obj.timePart = maybeTime;
        if (!obj.dateTime) obj.dateTime = dt;
      } else if (h.includes('เวลา') || lowerH.includes('time')) {
        const t = formatTimeOnly(val);
        if (t) obj.timePart = t;
      } else if (lowerH.includes('username') || lowerH === 'user' || lowerH.includes('user id') || h.includes('ยูสเซอร์')) {
        obj.username = String(val ?? '').trim();
      } else if (hasJakColumn ? h === 'จาก' : (h.includes('ธนาคาร') || lowerH.includes('bank'))) {
        obj.bank = String(val ?? '').trim();
      } else if (h.includes('เลขบัญชี') || h.includes('เลขที่บัญชี') || lowerH.includes('acc')) {
        obj.accNo = formatAccountNo(val);
      } else if (
        !h.includes('คงเหลือ') &&
        !lowerH.includes('balance') &&
        (
          h.includes('ยอดฝาก') ||
          h.includes('ยอดถอน') ||
          h.includes('จำนวนเงิน') ||
          h.includes('ยอดโอน') ||
          h.includes('ยอดเงิน') ||
          lowerH.includes('deposit') ||
          lowerH.includes('withdraw') ||
          lowerH.includes('amount')
        )
      ) {
        const parsed = parseNumeric(val);
        if (parsed > 0 || obj.amount == null) obj.amount = parsed;
      }
    });

    if (obj.datePart) {
      obj.dateTime = obj.timePart ? `${obj.datePart} ${obj.timePart}` : obj.datePart;
    }
    return obj;
  }).filter((r) => (r.dateTime || r.username) && Number(r.amount) > 0);

  const parsed = rawRows.map((raw) => {
    const { dateStr, timeStr } = extractDateAndTime(raw.dateTime);
    const txType = isWithdrawFile ? 'Withdraw' : 'Deposit';
    return [dateStr, raw.bank || 'Auto', raw.accNo || '0000', txType, raw.username || 'user', timeStr, Number((raw.amount || 0).toFixed(2))];
  });

  parsed.sort((a, b) => parseSheetDateTimeMs(a[0], a[5]) - parseSheetDateTimeMs(b[0], b[5]));
  return parsed;
}
