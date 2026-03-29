/**
 * Nettoyage des doublons dans le Google Sheet Sandro Temps.
 *
 * Usage:
 * 1) Ouvrir le Google Sheet
 * 2) Extensions > Apps Script
 * 3) Coller ce fichier dans un nouvel onglet de script
 * 4) Exécuter cleanupDuplicates()
 *
 * Le script:
 * - Supprime les lignes vides
 * - Supprime les doublons (même Date + Code Affaire + Durée)
 *   en gardant la ligne avec la description la plus longue
 * - Crée un log dans une popup avec le résumé
 */

function cleanupDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSourceSheet_(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Aucun onglet source trouvé (colonnes attendues: Date, Affaire, Durée (h)).');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('Aucune donnée à analyser.');
    return;
  }

  const headers = data[0].map(function(h) { return String(h || '').trim(); });
  const dateCol = headers.indexOf('Date');
  const codeCol = headers.indexOf('Code Affaire');
  const durationCol = headers.indexOf('Durée (h)');
  const descCol = headers.indexOf('Description');
  const statutCol = headers.indexOf('Statut');

  if (dateCol === -1 || codeCol === -1 || durationCol === -1) {
    SpreadsheetApp.getUi().alert('Colonnes introuvables: Date / Code Affaire / Durée (h).');
    return;
  }

  var emptyCount = 0;
  var duplicateCount = 0;
  var keptCount = 0;
  var skippedDeleted = 0;
  var seen = {}; // key -> { rowIndex, descLength }
  var rowsToDelete = []; // 0-based indices of rows to delete

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Ignorer les lignes supprimées (soft delete) — ne pas les compter comme doublons
    if (statutCol !== -1) {
      var statut = String(row[statutCol] || '').trim().toLowerCase();
      if (statut === 'supprimé' || statut === 'supprime') {
        skippedDeleted++;
        continue;
      }
    }

    var dateVal = normalizeDate_(row[dateCol]);
    var codeVal = String(row[codeCol] || '').trim();
    var durationVal = normalizeDuration_(row[durationCol]);
    var descVal = String(row[descCol] || '').trim();

    // Skip empty rows (no date or no code affaire)
    if (!dateVal || !codeVal) {
      emptyCount++;
      rowsToDelete.push(i);
      continue;
    }

    // Deduplication key: Date + Code Affaire + Duration
    var key = dateVal + '|' + codeVal + '|' + durationVal;

    if (seen[key] !== undefined) {
      // Duplicate found - keep the one with the longest description
      var existing = seen[key];
      if (descVal.length > existing.descLength) {
        // New row has longer description, delete the old one
        rowsToDelete.push(existing.rowIndex);
        seen[key] = { rowIndex: i, descLength: descVal.length };
      } else {
        // Old row has longer or equal description, delete the new one
        rowsToDelete.push(i);
      }
      duplicateCount++;
    } else {
      seen[key] = { rowIndex: i, descLength: descVal.length };
      keptCount++;
    }
  }

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getUi().alert('Aucun doublon ni ligne vide trouvé. Le sheet est propre.');
    return;
  }

  // Confirmation before deleting
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Nettoyage Sandro Temps',
    'Lignes analysées: ' + (data.length - 1) + '\n' +
    'Lignes supprimées (soft delete) ignorées: ' + skippedDeleted + '\n' +
    'Lignes vides à supprimer: ' + emptyCount + '\n' +
    'Doublons à supprimer: ' + duplicateCount + '\n' +
    'Lignes conservées: ' + keptCount + '\n\n' +
    'Total lignes à supprimer: ' + rowsToDelete.length + '\n\n' +
    'Confirmer la suppression ?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Opération annulée.');
    return;
  }

  // Delete rows from bottom to top to preserve indices
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j] + 1); // +1 because sheet rows are 1-based
  }

  SpreadsheetApp.flush();
  ss.toast(
    rowsToDelete.length + ' lignes supprimées (' + emptyCount + ' vides, ' + duplicateCount + ' doublons). ' + keptCount + ' lignes conservées.',
    'Nettoyage terminé',
    10
  );
}

function normalizeDate_(value) {
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var d = String(value.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  var str = String(value || '').trim();
  // ISO format
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return str;
  // FR format dd/mm/yyyy
  var frMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (frMatch) return frMatch[3] + '-' + frMatch[2].padStart(2, '0') + '-' + frMatch[1].padStart(2, '0');
  return '';
}

function normalizeDuration_(value) {
  if (typeof value === 'number') return String(Math.round(value * 100) / 100);
  var str = String(value || '').trim().replace(',', '.');
  var num = parseFloat(str);
  if (isNaN(num)) return '0';
  return String(Math.round(num * 100) / 100);
}

function findSourceSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (s.getName() === 'Dashboard Sandro') continue;
    var width = Math.max(1, s.getLastColumn());
    var header = s.getRange(1, 1, 1, width).getValues()[0].map(function(v) { return String(v || '').trim(); });
    if (header.indexOf('Date') !== -1 && header.indexOf('Code Affaire') !== -1 && header.indexOf('Durée (h)') !== -1) {
      return s;
    }
  }
  return null;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sandro Outils')
    .addItem('Nettoyer les doublons', 'cleanupDuplicates')
    .addToUi();
}
