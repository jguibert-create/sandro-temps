/**
 * Google Apps Script d'ingestion pour Sandro Temps.
 *
 * Ce script gère la synchronisation entre l'application web et le Google Sheet.
 *
 * Colonnes attendues dans le Sheet :
 *   Timestamp | Date | Code Affaire | Affaire | Durée (h) | Description | Entry ID | Statut | Semaine | Mois
 *
 * Usage :
 * 1) Ouvrir le Google Sheet
 * 2) Extensions > Apps Script
 * 3) Remplacer le code existant par ce fichier
 * 4) Déployer > Nouveau déploiement > Application Web
 *    - Exécuter en tant que : Moi
 *    - Accès : Tout le monde
 * 5) Copier l'URL de déploiement dans l'app (Config > URL Google Apps Script)
 */

// ==========================================
// CONFIGURATION
// ==========================================

var EXPECTED_HEADERS = ['Timestamp', 'Date', 'Code Affaire', 'Affaire', 'Durée (h)', 'Description', 'Entry ID', 'Statut', 'Semaine', 'Mois'];

// ==========================================
// ENTRY POINTS
// ==========================================

function doGet(e) {
  var params = e ? e.parameter : {};

  // Healthcheck
  if (params.test === '1') {
    return jsonResponse({ status: 'OK', message: 'Sandro Temps API active' });
  }

  // Export all active entries
  if (params.action === 'export') {
    return handleExport();
  }

  return jsonResponse({ status: 'OK', message: 'Sandro Temps API active' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || 'upsert';

    if (action === 'delete') {
      return handleDelete(payload);
    }

    return handleUpsert(payload);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Erreur: ' + err.message });
  }
}

// ==========================================
// HANDLERS
// ==========================================

function handleUpsert(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSourceSheet(ss);
  if (!sheet) {
    return jsonResponse({ success: false, message: 'Onglet source introuvable' });
  }

  var cols = getColumnMap(sheet);
  var entryId = String(payload.entryId || '');
  if (!entryId) {
    return jsonResponse({ success: false, message: 'entryId manquant' });
  }

  // Chercher si l'entrée existe déjà
  var existingRow = findRowByEntryId(sheet, cols, entryId);

  var now = new Date();
  var dateStr = String(payload.date || '');
  var weekNum = dateStr ? getWeekNumber(dateStr) : '';
  var monthLabel = dateStr ? getMonthLabel(dateStr) : '';

  var rowData = buildRowArray(cols, {
    timestamp: Utilities.formatDate(now, 'Europe/Paris', 'dd/MM/yyyy'),
    date: dateStr,
    codeAffaire: String(payload.affaireCode || ''),
    affaire: String(payload.affaireLabel || ''),
    duration: payload.duration,
    description: String(payload.description || ''),
    entryId: entryId,
    statut: 'actif',
    semaine: weekNum,
    mois: monthLabel
  });

  if (existingRow > 0) {
    // Update existing row
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    return jsonResponse({ success: true, action: 'updated', entryId: entryId });
  } else {
    // Append new row
    sheet.appendRow(rowData);
    return jsonResponse({ success: true, action: 'created', entryId: entryId });
  }
}

function handleDelete(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSourceSheet(ss);
  if (!sheet) {
    return jsonResponse({ success: false, message: 'Onglet source introuvable' });
  }

  var cols = getColumnMap(sheet);
  var entryId = String(payload.entryId || '');
  if (!entryId) {
    return jsonResponse({ success: false, message: 'entryId manquant' });
  }

  var existingRow = findRowByEntryId(sheet, cols, entryId);
  if (existingRow > 0) {
    if (cols.statut !== -1) {
      sheet.getRange(existingRow, cols.statut + 1).setValue('supprimé');
    }
    return jsonResponse({ success: true, action: 'deleted', entryId: entryId });
  }

  // Entry not found — consider it already deleted
  return jsonResponse({ success: true, action: 'not_found', entryId: entryId });
}

function handleExport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSourceSheet(ss);
  if (!sheet) {
    return jsonResponse({ success: false, message: 'Onglet source introuvable' });
  }

  var cols = getColumnMap(sheet);
  var data = sheet.getDataRange().getValues();
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Paris';
  var entries = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Filtrer les lignes supprimées
    if (cols.statut !== -1) {
      var statut = String(row[cols.statut] || '').trim().toLowerCase();
      if (statut === 'supprimé' || statut === 'supprime') continue;
    }

    var dateKey = toDateKey(row[cols.date], tz);
    var codeAffaire = String(row[cols.codeAffaire] || '').trim();
    var duration = toNumber(row[cols.duration]);

    if (!dateKey || !codeAffaire || !isFinite(duration) || duration <= 0) continue;

    var entryId = '';
    if (cols.entryId !== -1) {
      entryId = String(row[cols.entryId] || '').trim();
    }
    // Fallback : synthétiser un ID si absent
    if (!entryId) {
      entryId = dateKey + '|' + codeAffaire + '|' + duration + '|' + String(row[cols.description] || '').trim();
    }

    var entry = {
      id: entryId,
      date: dateKey,
      affaire: codeAffaire,
      duration: duration,
      description: String(row[cols.description] || '').trim(),
      autoFill: false
    };

    if (!entries[dateKey]) entries[dateKey] = [];
    entries[dateKey].push(entry);
  }

  return jsonResponse({ success: true, entries: entries });
}

// ==========================================
// HELPERS
// ==========================================

function findSourceSheet(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var name = s.getName().toLowerCase();
    if (name === 'dashboard sandro') continue;
    var width = Math.max(1, s.getLastColumn());
    if (width < 3) continue;
    var header = s.getRange(1, 1, 1, width).getValues()[0].map(function(v) { return String(v || '').trim(); });
    if (header.indexOf('Date') !== -1 && header.indexOf('Durée (h)') !== -1) {
      return s;
    }
  }
  return null;
}

function getColumnMap(sheet) {
  var width = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(function(v) { return String(v || '').trim(); });
  return {
    timestamp: headers.indexOf('Timestamp'),
    date: headers.indexOf('Date'),
    codeAffaire: headers.indexOf('Code Affaire'),
    affaire: headers.indexOf('Affaire'),
    duration: headers.indexOf('Durée (h)'),
    description: headers.indexOf('Description'),
    entryId: headers.indexOf('Entry ID'),
    statut: headers.indexOf('Statut'),
    semaine: headers.indexOf('Semaine'),
    mois: headers.indexOf('Mois'),
    count: headers.length
  };
}

function findRowByEntryId(sheet, cols, entryId) {
  if (cols.entryId === -1) return -1;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var idColumn = sheet.getRange(2, cols.entryId + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idColumn.length; i++) {
    if (String(idColumn[i][0] || '').trim() === entryId) {
      return i + 2; // +2 car row 1 = header, i est 0-based
    }
  }
  return -1;
}

function buildRowArray(cols, data) {
  var row = [];
  for (var c = 0; c < cols.count; c++) {
    if (c === cols.timestamp) row.push(data.timestamp);
    else if (c === cols.date) row.push(data.date);
    else if (c === cols.codeAffaire) row.push(data.codeAffaire);
    else if (c === cols.affaire) row.push(data.affaire);
    else if (c === cols.duration) row.push(data.duration);
    else if (c === cols.description) row.push(data.description);
    else if (c === cols.entryId) row.push(data.entryId);
    else if (c === cols.statut) row.push(data.statut);
    else if (c === cols.semaine) row.push(data.semaine);
    else if (c === cols.mois) row.push(data.mois);
    else row.push('');
  }
  return row;
}

function toDateKey(value, tz) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  var str = String(value || '').trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return str;
  var fr = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) return fr[3] + '-' + fr[2].padStart(2, '0') + '-' + fr[1].padStart(2, '0');
  return '';
}

function toNumber(value) {
  if (typeof value === 'number') return value;
  var str = String(value || '').trim().replace(',', '.');
  if (!str) return NaN;
  return Number(str);
}

function getWeekNumber(dateStr) {
  try {
    var parts = dateStr.split('-');
    var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    var dayOfYear = Math.floor((d - new Date(d.getFullYear(), 0, 1)) / 86400000);
    var weekNum = Math.ceil((dayOfYear + new Date(d.getFullYear(), 0, 1).getDay() + 1) / 7);
    return 'S' + weekNum;
  } catch (e) {
    return '';
  }
}

function getMonthLabel(dateStr) {
  try {
    var parts = dateStr.split('-');
    var months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin',
                  'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
    var monthIdx = Number(parts[1]) - 1;
    return months[monthIdx] + ' ' + parts[0];
  } catch (e) {
    return '';
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
