// ═══════════════════════════════════════════════
// BİASS SİPARİŞ UYGULAMASI — APPS SCRIPT
// Google Sheet ID'yi aşağıya girin
// ═══════════════════════════════════════════════
const SHEET_ID = 'BURAYA_SHEET_ID_YAZIN';

// Kullanıcı mail adresleri
const USER_EMAILS = {
  'oguzhan':  'oguzhan@biass.com.tr',
  'hamza':    'hamzatajimuradow@gmail.com',
  'akif':     'akifarici@biass.com.tr',
  'ahmet':    'ahmetarici@biass.com.tr',
  'cem':      'depo@biass.com.tr',
  'mersan':   'mersanyildirim@biass.com.tr',
  'yasin':    'yasinsuslu@biass.com.tr',
  'yonetici': '',
};

// Fabrika sorumluları
const FABRIKA_USERS = ['oguzhan', 'hamza'];

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreate(ss, 'SiparisDB');
    const val = sheet.getRange(1, 1).getValue();
    return ContentService
      .createTextOutput(val || '{"orders":[],"users":[],"version":1}')
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreate(ss, 'SiparisDB');
    const payload = JSON.parse(e.postData.contents);

    // Sipariş ekle
    if (payload.action === 'addOrder') {
      const current = sheet.getRange(1,1).getValue();
      const data = current ? JSON.parse(current) : {orders:[], users:[]};
      data.orders = data.orders || [];
      const exists = data.orders.some(o => o.id === payload.order.id);
      if (!exists) {
        data.orders.push(payload.order);
        sheet.getRange(1,1).setValue(JSON.stringify(data));
        updateOrderSheet(ss, data.orders);
        // Fabrika sorumlularına mail gönder
        sendOrderNotification(payload.order, 'new');
      }
      return ok();
    }

    // Sipariş güncelle (durum değişikliği)
    if (payload.action === 'updateOrder') {
      const current = sheet.getRange(1,1).getValue();
      const data = current ? JSON.parse(current) : {orders:[]};
      data.orders = data.orders || [];
      const idx = data.orders.findIndex(o => o.id === payload.order.id);
      if (idx >= 0) {
        data.orders[idx] = payload.order;
        sheet.getRange(1,1).setValue(JSON.stringify(data));
        updateOrderSheet(ss, data.orders);
        // Sipariş sahibine mail gönder
        sendStatusNotification(payload.order, payload.prevStatus);
      }
      return ok();
    }

    // Tam veri yazma (kullanıcılar vs)
    const current = sheet.getRange(1,1).getValue();
    const existing = current ? JSON.parse(current) : {};
    const merged = Object.assign(existing, payload);
    sheet.getRange(1,1).setValue(JSON.stringify(merged));

    return ok();
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ok() {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendOrderNotification(order, type) {
  // Fabrika sorumlularına yeni sipariş bildirimi
  const subject = '🔔 Yeni Sipariş: ' + order.id;
  const urunler = (order.items || []).map(i => '• ' + formatItem(i)).join('\n');
  const body = `Merhaba,

Yeni bir sipariş oluşturuldu.

Sipariş No: ${order.id}
Oluşturan: ${order.createdByName}
Tarih: ${order.createdAt}
Müşteri Notu: ${order.note || '—'}

Ürünler:
${urunler}

Lütfen uygulamadan siparişi onaylayın ve tahmini termin bilgisini girin.

Biass Sipariş Sistemi`;

  FABRIKA_USERS.forEach(uid => {
    const email = USER_EMAILS[uid];
    if (email) {
      try { MailApp.sendEmail(email, subject, body); } catch(e) {}
    }
  });
}

function sendStatusNotification(order, prevStatus) {
  const ownerEmail = USER_EMAILS[order.createdBy];
  if (!ownerEmail) return;

  const statusLabels = {
    'alinmadi': 'Sipariş Alınmadı',
    'alindi': 'Sipariş Alındı',
    'islemde': 'Sipariş İşlemde',
    'hazir': 'Sipariş Hazır'
  };

  const newLabel = statusLabels[order.status] || order.status;
  const subject = '📦 Sipariş Durumu Güncellendi: ' + newLabel;
  const urunler = (order.items || []).map(i => '• ' + formatItem(i)).join('\n');

  let extra = '';
  if (order.status === 'alindi' && order.termin) {
    extra = `\nTahmini Termin: ${order.termin}`;
  }
  if (order.status === 'hazir') {
    extra = '\n✅ Siparişiniz hazır! Teslim alabilirsiniz.';
  }

  const body = `Merhaba ${order.createdByName},

Siparişinizin durumu güncellendi.

Sipariş No: ${order.id}
Yeni Durum: ${newLabel}${extra}

Ürünler:
${urunler}

Biass Sipariş Sistemi`;

  try { MailApp.sendEmail(ownerEmail, subject, body); } catch(e) {}
}

function formatItem(item) {
  if (!item) return '?';
  let s = item.category + ': ';
  if (item.productCode) s += item.productCode + ' ';
  if (item.size) s += item.size + ' ';
  if (item.color) s += '/ ' + item.color + ' ';
  if (item.packaging) s += '/ ' + item.packaging + ' ';
  if (item.quantity) s += '— ' + item.quantity;
  if (item.note) s += ' (Not: ' + item.note + ')';
  return s;
}

function updateOrderSheet(ss, orders) {
  if (!orders || !orders.length) return;
  const logSheet = getOrCreate(ss, 'Siparişler');
  logSheet.clearContents();
  logSheet.getRange(1,1,1,8).setValues([
    ['Sipariş No','Oluşturan','Tarih','Durum','Termin','Ürün Sayısı','Not','Son Güncelleme']
  ]);
  const rows = orders.map(o => [
    o.id,
    o.createdByName || '',
    o.createdAt || '',
    o.status || '',
    o.termin || '',
    (o.items || []).length,
    o.note || '',
    o.updatedAt || ''
  ]);
  if (rows.length) logSheet.getRange(2,1,rows.length,8).setValues(rows);
}

// Zamanlanmış hatırlatıcı — her saat çalıştır
// Apps Script → Tetikleyiciler → addReminderTrigger
function checkPendingOrders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SiparisDB');
  if (!sheet) return;
  const val = sheet.getRange(1,1).getValue();
  if (!val) return;
  const data = JSON.parse(val);
  const pending = (data.orders || []).filter(o => o.status === 'alinmadi');
  if (!pending.length) return;

  pending.forEach(order => {
    const subject = '⚠️ Bekleyen Sipariş: ' + order.id;
    const body = `Merhaba,

Bu sipariş henüz alınmadı olarak işaretlenmedi.

Sipariş No: ${order.id}
Oluşturan: ${order.createdByName}
Tarih: ${order.createdAt}

Lütfen uygulamadan siparişi onaylayın.

Biass Sipariş Sistemi`;

    FABRIKA_USERS.forEach(uid => {
      const email = USER_EMAILS[uid];
      if (email) {
        try { MailApp.sendEmail(email, subject, body); } catch(e) {}
      }
    });
  });
}
