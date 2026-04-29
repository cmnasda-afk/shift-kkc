// SHIFT KKC — Google Sheets Sync Backend v3
// วางโค้ดนี้ใน Google Apps Script (Extensions → Apps Script)
// Deploy as Web App → Execute as: Me, Who has access: Anyone

const SS = SpreadsheetApp.getActiveSpreadsheet();

// headers ต้องตรงกัน — เพิ่ม staffName ตอนท้าย
const ENTRY_HEADERS = ['id','type','date','ts','name','unit','qty','status','expDate',
  'recvDate','supplier','note','priceTotal','pricePerUnit','updatedAt','staffName'];

const BALANCE_HEADERS = ['name','unit','qty','status','lastCheckDate','lastUpdate','lastStaff'];

function getOrCreateSheet(name, headers) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // auto-extend headers ถ้า column ใหม่ถูกเพิ่มในโค้ด
  const lastCol = sh.getLastColumn();
  if (lastCol < headers.length) {
    sh.getRange(1, lastCol + 1, 1, headers.length - lastCol)
      .setValues([headers.slice(lastCol)]).setFontWeight('bold');
  }
  return sh;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'saveEntry') return saveEntry(body.entry);
    if (action === 'saveReceive') return saveReceive(body.receive);
    if (action === 'saveCustomItem') return saveCustomItem(body.item);
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getBalance') return getBalance();
    if (action === 'getCustomItems') return getCustomItems();
    if (action === 'getEntries') return getEntries(e.parameter.since);
    if (action === 'getAll') return getAll();
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMonthSheetName(dateStr) {
  if (!dateStr) {
    const now = new Date();
    return Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
  }
  return dateStr.substring(0, 7);
}

function entryRow(entry) {
  return [
    entry.id || '', entry.type || 'stock', entry.date || '',
    entry.ts || Date.now(), entry.name || '', entry.unit || '',
    entry.qty || '', entry.status || '', entry.expDate || '',
    entry.recvDate || '', entry.supplier || '', entry.note || '',
    entry.priceTotal || '', entry.pricePerUnit || '',
    new Date().toISOString(), entry.staffName || ''
  ];
}

function saveEntry(entry) {
  if (!entry || !entry.name) return out({ ok: false, error: 'no entry' });

  const monthName = getMonthSheetName(entry.date);
  const monthSh = getOrCreateSheet(monthName, ENTRY_HEADERS);
  monthSh.appendRow(entryRow(entry));

  updateBalance(entry);

  const sh = getOrCreateSheet('Entries', ENTRY_HEADERS);
  sh.appendRow(entryRow(entry));

  return out({ ok: true });
}

function updateBalance(entry) {
  if (!entry.name || entry.type === 'receive') return;

  const sh = getOrCreateSheet('Balance', BALANCE_HEADERS);
  const data = sh.getDataRange().getValues();

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === entry.name) { foundRow = i + 1; break; }
  }

  const rowData = [
    entry.name, entry.unit || '', entry.qty || '',
    entry.status || 'ok', entry.date || '',
    new Date().toISOString(), entry.staffName || ''
  ];

  if (foundRow > 0) {
    sh.getRange(foundRow, 1, 1, BALANCE_HEADERS.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
}

function saveReceive(receive) {
  return saveEntry(receive);
}

function saveCustomItem(item) {
  if (!item || !item.name) return out({ ok: false, error: 'no item' });
  const sh = getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === item.name && data[i][0] === (item.cat || '')) return out({ ok: true });
  }
  sh.appendRow([item.cat || '', item.name, item.unit || '', new Date().toISOString()]);
  return out({ ok: true });
}

function getBalance() {
  const sh = SS.getSheetByName('Balance');
  if (!sh) return getBalanceFallback();

  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return getBalanceFallback();

  const balance = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const name = r[0];
    const qty = r[2];
    if (!name || qty === '' || qty === undefined || qty === null) continue;
    balance.push({
      name: name,
      unit: r[1] || '',
      currentQty: String(qty),
      currentStatus: r[3] || 'ok',
      lastCheckDate: r[4] || '',
      updatedAt: r[5] || '',
      lastStaff: r[6] || ''
    });
  }

  return out({ ok: true, balance: balance });
}

function getBalanceFallback() {
  const sh = SS.getSheetByName('Entries');
  if (!sh) return out({ ok: true, balance: [] });
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return out({ ok: true, balance: [] });
  const headers = data[0];
  const rows = data.slice(1);
  const idx = (h) => headers.indexOf(h);

  const byName = {};
  rows.forEach(r => {
    const name = r[idx('name')];
    if (!name) return;
    const type = r[idx('type')];
    const ts = Number(r[idx('ts')]) || 0;
    if (!byName[name]) byName[name] = { name: name, stockTs: 0, recvTs: 0 };
    const rec = byName[name];
    if (type === 'stock' || !rec.currentQty) {
      if (ts > (rec.stockTs || 0)) {
        rec.stockTs = ts;
        rec.unit = r[idx('unit')];
        rec.currentQty = String(r[idx('qty')] || '');
        rec.currentStatus = r[idx('status')] || 'ok';
        rec.lastCheckDate = r[idx('date')] || '';
        rec.updatedAt = r[idx('updatedAt')] || '';
        rec.lastStaff = idx('staffName') >= 0 ? (r[idx('staffName')] || '') : '';
      }
    }
  });

  const balance = Object.values(byName).filter(b => b.currentQty !== undefined && b.currentQty !== '');
  return out({ ok: true, balance: balance });
}

// ─── Get latest entries (newest first), optional since=timestamp ───
function getEntries(since) {
  const sh = SS.getSheetByName('Entries');
  if (!sh) return out({ ok: true, entries: [] });
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return out({ ok: true, entries: [] });

  const headers = data[0];
  const idx = (h) => headers.indexOf(h);
  const sinceTs = since ? Number(since) : 0;

  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const ts = Number(r[idx('ts')]) || 0;
    if (sinceTs && ts <= sinceTs) continue;
    const name = r[idx('name')];
    if (!name) continue;
    entries.push({
      id: r[idx('id')] || '',
      type: r[idx('type')] || 'stock',
      date: r[idx('date')] || '',
      ts: ts,
      name: name,
      unit: r[idx('unit')] || '',
      qty: String(r[idx('qty')] || ''),
      status: r[idx('status')] || 'ok',
      expDate: r[idx('expDate')] || '',
      recvDate: r[idx('recvDate')] || '',
      supplier: r[idx('supplier')] || '',
      note: r[idx('note')] || '',
      priceTotal: r[idx('priceTotal')] || '',
      pricePerUnit: r[idx('pricePerUnit')] || '',
      staffName: idx('staffName') >= 0 ? (r[idx('staffName')] || '') : ''
    });
  }
  // newest first, cap 500
  entries.sort((a,b) => b.ts - a.ts);
  return out({ ok: true, entries: entries.slice(0, 500) });
}

// ─── Get balance + customItems in one call (ลด round-trip) ───
function getAll() {
  const balRes = JSON.parse(getBalance().getContent());
  const itemsRes = JSON.parse(getCustomItems().getContent());
  return out({
    ok: true,
    balance: balRes.balance || [],
    items: itemsRes.items || [],
    serverTs: Date.now()
  });
}

function getCustomItems() {
  const sh = SS.getSheetByName('CustomItems');
  if (!sh) return out({ ok: true, items: [] });
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return out({ ok: true, items: [] });
  const items = data.slice(1)
    .filter(r => r[1])
    .map(r => ({ cat: r[0], name: r[1], unit: r[2] }));
  return out({ ok: true, items });
}

// ─── Test function (run manually from editor) ───
function test_setup() {
  getOrCreateSheet('Entries', ENTRY_HEADERS);
  getOrCreateSheet('Balance', BALANCE_HEADERS);
  getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  const monthName = getMonthSheetName();
  getOrCreateSheet(monthName, ENTRY_HEADERS);
  Logger.log('Setup complete. Sheets: Entries, Balance, CustomItems, ' + monthName);
}
