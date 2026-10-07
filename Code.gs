/**
 * Créathèque — API Google Sheets + Drive
 * À coller dans : Google Sheet > Extensions > Apps Script
 *
 * 1. Change MOT_DE_PASSE ci-dessous (c'est lui que tu tapes dans la page).
 * 2. Lance la fonction `setup` une fois (bouton ▶ Exécuter).
 * 3. Déployer > Nouveau déploiement > Application Web
 *    Exécuter en tant que : Moi — Qui a accès : Tout le monde.
 */

const MOT_DE_PASSE = 'change-moi';

const SHEET = 'Créas';
const COLS = ['id','name','status','boutique','produit','angle','format','hook','notes',
  'fileId','mime','createdAt','launchedAt','endedAt','spend','roas','ctr','cpa',
  'parentId','parentName','updatedAt'];
const DATE_COLS = ['createdAt','launchedAt','endedAt','updatedAt'];
const NUM_COLS = ['spend','roas','ctr','cpa'];
const FORMATS = COLS.map(c => DATE_COLS.includes(c) ? 'dd/mm/yyyy hh:mm' : NUM_COLS.includes(c) ? '0.00' : '@');

/* ---------- installation (à lancer une fois) ---------- */
function setup() {
  if (MOT_DE_PASSE === 'change-moi') throw new Error('Change MOT_DE_PASSE en haut du script avant de lancer setup.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('Créathèque — visuels');
    props.setProperty('FOLDER_ID', f.getId());
  }
  Logger.log('OK. Dossier Drive : ' + folder_().getUrl());
}

/* ---------- API ---------- */
function doGet() {
  return json_({ ok: true, message: 'API Créathèque en ligne. Utilise la page pour y accéder.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (String(req.key) !== MOT_DE_PASSE) return json_({ ok: false, error: 'unauthorized' });
    switch (req.action) {
      case 'list':
        return json_({ ok: true, items: listAll_(),
          sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(), folderUrl: folder_().getUrl() });
      case 'save':
        return json_({ ok: true, item: save_(req.item) });
      case 'delete':
        delete_(req.id, req.deleteFile);
        return json_({ ok: true });
      case 'upload':
        return json_({ ok: true, fileId: upload_(req) });
      default:
        return json_({ ok: false, error: 'Action inconnue : ' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- données ---------- */
function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error('Onglet "' + SHEET + '" introuvable : lance setup.');
  return sh;
}

function folder_() {
  const id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!id) throw new Error('Dossier Drive non créé : lance setup.');
  return DriveApp.getFolderById(id);
}

function listAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, last - 1, head.length).getValues();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((h, i) => {
      let v = r[i];
      if (v === '' || v === null) v = null;
      else if (v instanceof Date) v = v.getTime();
      else if (NUM_COLS.includes(h)) v = Number(v);
      o[h] = v;
    });
    return o;
  });
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function save_(item) {
  if (!item || !item.id) throw new Error('Créa sans identifiant.');
  item.updatedAt = Date.now();
  const row = COLS.map(c => {
    const v = item[c];
    if (v === null || v === undefined || v === '') return '';
    if (DATE_COLS.includes(c)) return new Date(Number(v));
    if (NUM_COLS.includes(c)) return Number(v);
    return String(v);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    let r = findRow_(sh, item.id);
    if (r < 0) r = sh.getLastRow() + 1;
    const range = sh.getRange(r, 1, 1, COLS.length);
    range.setNumberFormats([FORMATS]);
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return item;
}

function delete_(id, deleteFile) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const r = findRow_(sh, id);
    if (r > 0) sh.deleteRow(r);
  } finally {
    lock.releaseLock();
  }
  if (deleteFile) {
    try { DriveApp.getFileById(deleteFile).setTrashed(true); } catch (e) {}
  }
}

function upload_(req) {
  const blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime, req.name);
  const file = folder_().createFile(blob);
  // Lien "toute personne disposant du lien" : nécessaire pour afficher les miniatures dans la page.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
