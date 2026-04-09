// SHIFT KKC — Google Sheets Sync Backend
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

// ─── Save stock/receive/loss entry ───
function saveEntry(entry) {
  if (!entry || !entry.name) return out({ ok: false, error: 'no entry' });
  const sh = getOrCreateSheet('Entries', [
    'id','type','date','ts','name','unit','qty','status','expDate',
    'recvDate','supplier','note','priceTotal','pricePerUnit','updatedAt'
  ]);
  sh.appendRow([
    entry.id || '',
    entry.type || 'stock',
    entry.date || '',
    entry.ts || Date.now(),
    entry.name || '',
    entry.unit || '',
    entry.qty || '',
    entry.status || '',
    entry.expDate || '',
    entry.recvDate || '',
    entry.supplier || '',
    entry.note || '',
    entry.priceTotal || '',
    entry.pricePerUnit || '',
    new Date().toISOString()
  ]);
  return out({ ok: true });
}

function saveReceive(receive) {
  return saveEntry(receive);
}

function saveCustomItem(item) {
  if (!item || !item.name) return out({ ok: false, error: 'no item' });
  const sh = getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  sh.appendRow([item.cat || '', item.name, item.unit || '', new Date().toISOString()]);
  return out({ ok: true });
}

// ─── Get latest balance per item ───
function getBalance() {
  const sh = SS.getSheetByName('Entries');
  if (!sh) return out({ ok: true, balance: [] });
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return out({ ok: true, balance: [] });
  const headers = data[0];
  const rows = data.slice(1);

  // map by item name: keep latest stock entry, plus latest receive info
  const byName = {};
  rows.forEach(r => {
    const type = r[headers.indexOf('type')];
    const name = r[headers.indexOf('name')];
    if (!name) return;
    const ts = Number(r[headers.indexOf('ts')]) || 0;
    if (!byName[name]) byName[name] = {};
    const rec = byName[name];

    if (type === 'stock') {
      if (!rec.stockTs || ts > rec.stockTs) {
        rec.stockTs = ts;
        rec.name = name;
        rec.unit = r[headers.indexOf('unit')];
        rec.currentQty = r[headers.indexOf('qty')];
        rec.currentStatus = r[headers.indexOf('status')];
        rec.expDate = r[headers.indexOf('expDate')];
        rec.note = r[headers.indexOf('note')];
        rec.lastCheckDate = r[headers.indexOf('date')];
        rec.updatedAt = r[headers.indexOf('updatedAt')];
      }
    } else if (type === 'receive') {
      if (!rec.recvTs || ts > rec.recvTs) {
        rec.recvTs = ts;
        rec.lastRecvDate = r[headers.indexOf('recvDate')] || r[headers.indexOf('date')];
        rec.lastRecvSupplier = r[headers.indexOf('supplier')];
      }
    }
  });

  const balance = Object.values(byName).filter(b => b.currentQty !== undefined);
  return out({ ok: true, balance });
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
  getOrCreateSheet('CustomItems', ['cat','name','unit','addedAt']);
  Logger.log('Setup complete. Sheets created.');
}
