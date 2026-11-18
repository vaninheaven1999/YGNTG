import * as XLSX from 'xlsx';

export const LEDGER_HEADERS = ['Date', 'Bank', 'Acc No', 'Type', 'Username', 'Time', 'Amount', 'Remark'];

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

function formatExcelDateTime(val) {
  if (val instanceof Date) {
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

export function extractDateAndTime(dtStr, dateFormatPref = 'D-M-YYYY') {
  if (!dtStr) return { dateStr: '10-1-2026', timeStr: '0:00:00', timeSeconds: 0 };
  const parts = String(dtStr).trim().split(/\s+/);
  const datePart = parts[0] || '';
  const timePart = parts[1] || '00:00:00';
  let day = 1, month = 10, year = 2026;
  if (datePart.includes('/')) {
    const dp = datePart.split('/');
    day = parseInt(dp[0], 10) || 1; month = parseInt(dp[1], 10) || 10; year = parseInt(dp[2], 10) || 2026;
  }
  const formattedDate = `${day}-${month}-${year}`;
  const tp = timePart.split(':');
  const hr = parseInt(tp[0] || '0', 10), min = parseInt(tp[1] || '0', 10), sec = parseInt(tp[2] || '0', 10);
  return { dateStr: formattedDate, timeStr: `${hr}:${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')}`, timeSeconds: hr * 3600 + min * 60 + sec };
}

export function parseDepositExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) return [];
  const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', blankrows: false });
  if (!jsonData.length) return [];

  let headerIndex = jsonData.findIndex((row, i) => i < 10 && row.some((cell) => {
    if (typeof cell !== 'string') return false;
    const c = cell.toLowerCase();
    return cell.includes('วันที่') || c.includes('date') || c.includes('username') || cell.includes('ยอดฝาก') || cell.includes('ยอดถอน') || cell.includes('จำนวนเงิน') || cell.trim() === 'จาก';
  }));
  if (headerIndex === -1) headerIndex = 0;

  const headers = jsonData[headerIndex].map((h) => String(h || '').trim());
  const isWithdrawFile = headers.some((h) => h.includes('ยอดถอน') || h.toLowerCase().includes('withdraw'));
  const hasJakColumn = headers.some((h) => h === 'จาก');

  const rawRows = jsonData.slice(headerIndex + 1).map((row) => {
    const obj = {};
    headers.forEach((h, idx) => {
      if (!h) return;
      const val = row[idx];
      const lowerH = h.toLowerCase();

      if (h === 'วันที่' || lowerH.includes('date') || lowerH.includes('เวลา')) {
        obj.dateTime = formatExcelDateTime(val);
      } else if (lowerH.includes('username') || lowerH === 'user' || h.includes('ยูสเซอร์')) {
        obj.username = String(val ?? '').trim();
      } else if (hasJakColumn ? h === 'จาก' : (h.includes('ธนาคาร') || lowerH.includes('bank'))) {
        obj.bank = String(val ?? '').trim();
      } else if (h.includes('เลขบัญชี') || lowerH.includes('acc')) {
        obj.accNo = formatAccountNo(val);
      } else if (
        h.includes('ยอดฝาก') || h.includes('ยอดถอน') || h.includes('จำนวนเงิน') || h.includes('ยอดเงิน') ||
        lowerH.includes('deposit') || lowerH.includes('withdraw') || lowerH.includes('amount')
      ) {
        const parsed = parseNumeric(val);
        if (parsed > 0 || !obj.amount) obj.amount = parsed;
      }
    });
    return obj;
  }).filter((r) => r.dateTime || r.username || (r.amount && r.amount > 0));

  return rawRows.map((raw) => {
    const { dateStr, timeStr } = extractDateAndTime(raw.dateTime);
    const txType = isWithdrawFile ? 'Withdraw' : 'Deposit';
    return [dateStr, raw.bank || 'Auto', raw.accNo || '0000', txType, raw.username || 'user', timeStr, Number((raw.amount || 0).toFixed(2))];
  });
}
