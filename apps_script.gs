// SHIFT KKC — Google Sheets Sync Backend v2
// วางโค้ดนี้ใน Google Apps Script (Extensions → Apps Script)
// Deploy as Web App → Execute as: Me, Who has access: Anyone

const SS = SpreadsheetApp.getActiveSpreadsheet();

function getOrCreateSheet(name, headers) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
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
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Helper: get month sheet name ───
function getMonthSheetName(dateStr) {
  if (!dateStr) {
    const now = new Date();
    return Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
  }
  // dateStr = "2026-04-15" → "2026-04"
  return dateStr.substring(0, 7);
}

// ─── Save entry to monthly sheet + Balance sheet ───
function saveEntry(entry) {
  if (!entry || !entry.name) return out({ ok: false, error: 'no entry' });

  const headers = ['id','type','date','ts','name','unit','qty','status','expDate',
    'recvDate','supplier','note','priceTotal','pricePerUnit','updatedAt'];

  // 1. Save to monthly sheet (e.g., "2026-04")
  const monthName = getMonthSheetName(entry.date);
  const monthSh = getOrCreateSheet(monthName, headers);
  monthSh.appendRow([
    entry.id || '', entry.type || 'stock', entry.date || '',
    entry.ts || Date.now(), entry.name || '', entry.unit || '',
    entry.qty || '', entry.status || '', entry.expDate || '',
    entry.recvDate || '', entry.supplier || '', entry.note || '',
    entry.priceTotal || '', entry.pricePerUnit || '',
    new Date().toISOString()
  ]);

  // 2. Update Balance sheet (latest qty per item)
  updateBalance(entry);

  // 3. Also save to Entries (backward compat)
  const sh = getOrCreateSheet('Entries', headers);
  sh.appendRow([
    entry.id || '', entry.type || 'stock', entry.date || '',
    entry.ts || Date.now(), entry.name || '', entry.unit || '',
    entry.qty || '', entry.status || '', entry.expDate || '',
    entry.recvDate || '', entry.supplier || '', entry.note || '',
    entry.priceTotal || '', entry.pricePerUnit || '',
    new Date().toISOString()
  ]);

  return out({ ok: true });
}

// ─── Update Balance sheet (สรุปยอดล่าสุดต่อรายการ) ───
function updateBalance(entry) {
  if (!entry.name || entry.type === 'receive') return;

  const headers = ['name','unit','qty','status','lastCheckDate','lastUpdate'];
  const sh = getOrCreateSheet('Balance', headers);
  const data = sh.getDataRange().getValues();

  // find existing row for this item
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === entry.name) { foundRow = i + 1; break; }
  }

  const rowData = [
    entry.name, entry.unit || '', entry.qty || '',
    entry.status || 'ok', entry.date || '',
    new Date().toISOString()
  ];

  if (foundRow > 0) {
    // update existing row
    sh.getRange(foundRow, 1, 1, headers.length).setValues([rowData]);
  } else {
    // append new row
    sh.appendRow(rowData);
  }
}

function saveReceive(receive) {
  return saveEntry(receive);
}

function saveCustomItem(item) {
  if (!item || !item.name) return out({ ok: false, error: 'no item' });
  const sh = getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  // check duplicate
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === item.name && data[i][0] === (item.cat || '')) return out({ ok: true });
  }
  sh.appendRow([item.cat || '', item.name, item.unit || '', new Date().toISOString()]);
  return out({ ok: true });
}

// ─── Get latest balance per item (from Balance sheet) ───
function getBalance() {
  const sh = SS.getSheetByName('Balance');
  if (!sh) return getBalanceFallback(); // fallback to old Entries method

  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return getBalanceFallback();

  const headers = data[0];
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
      updatedAt: r[5] || ''
    });
  }

  return out({ ok: true, balance: balance });
}

// fallback: read from Entries sheet (old method)
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
      }
    }
  });

  const balance = Object.values(byName).filter(b => b.currentQty !== undefined && b.currentQty !== '');
  return out({ ok: true, balance: balance });
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
  getOrCreateSheet('Entries', [
    'id','type','date','ts','name','unit','qty','status','expDate',
    'recvDate','supplier','note','priceTotal','pricePerUnit','updatedAt'
  ]);
  getOrCreateSheet('Balance', ['name','unit','qty','status','lastCheckDate','lastUpdate']);
  getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  // create current month sheet
  const monthName = getMonthSheetName();
  getOrCreateSheet(monthName, [
    'id','type','date','ts','name','unit','qty','status','expDate',
    'recvDate','supplier','note','priceTotal','pricePerUnit','updatedAt'
  ]);
  Logger.log('Setup complete. Sheets created: Entries, Balance, CustomItems, ' + monthName);
}
