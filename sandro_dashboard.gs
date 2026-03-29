/**
 * Dashboard automatique pour le suivi Sandro.
 *
 * Usage:
 * 1) Ouvrir le Google Sheet
 * 2) Extensions > Apps Script
 * 3) Coller ce fichier et enregistrer
 * 4) Exécuter createOrUpdateSandroDashboard()
 *
 * Ensuite, utiliser le menu "Sandro Dashboard" dans le Sheet.
 */

const DASHBOARD_SHEET_NAME = 'Dashboard Sandro';
const DAILY_TARGET_HOURS = 8;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sandro Dashboard')
    .addItem('Créer / Actualiser le dashboard', 'createOrUpdateSandroDashboard')
    .addToUi();
}

function createOrUpdateSandroDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = findSourceSheet_(ss);
  if (!sourceSheet) {
    throw new Error('Aucun onglet source trouvé (colonnes attendues: Date, Affaire, Durée (h)).');
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length < 2) {
    throw new Error('Aucune donnée à analyser dans l’onglet source.');
  }

  const headers = sourceValues[0].map((h) => String(h || '').trim());
  const dateCol = headers.indexOf('Date');
  const affairCol = headers.indexOf('Affaire');
  const durationCol = headers.indexOf('Durée (h)');
  const timestampCol = headers.indexOf('Timestamp');
  if (dateCol === -1 || affairCol === -1 || durationCol === -1) {
    throw new Error('Colonnes introuvables: Date / Affaire / Durée (h).');
  }

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/Paris';
  const statutCol = headers.indexOf('Statut');
  const dailyTotals = new Map();
  const dailyByAffair = new Map();
  const monthlyTotals = new Map();
  const affairTotals = new Map();

  for (let i = 1; i < sourceValues.length; i++) {
    const row = sourceValues[i];

    // Ignorer les lignes supprimées (soft delete)
    if (statutCol !== -1) {
      const statut = String(row[statutCol] || '').trim().toLowerCase();
      if (statut === 'supprimé' || statut === 'supprime') continue;
    }

    const dateKey = toDateKey_(row[dateCol], tz, row[timestampCol]);
    const duration = toNumber_(row[durationCol]);
    if (!dateKey || !Number.isFinite(duration) || duration <= 0) continue;

    const affair = String(row[affairCol] || 'Sans affaire').trim() || 'Sans affaire';
    incrementMap_(dailyTotals, dateKey, duration);
    incrementMap_(monthlyTotals, dateKey.slice(0, 7), duration);
    incrementMap_(affairTotals, affair, duration);

    if (!dailyByAffair.has(dateKey)) dailyByAffair.set(dateKey, new Map());
    incrementMap_(dailyByAffair.get(dateKey), affair, duration);
  }

  const allDates = [...dailyTotals.keys()].sort();
  const allMonths = [...monthlyTotals.keys()].sort();
  const allAffairs = [...affairTotals.keys()].sort((a, b) => (affairTotals.get(b) - affairTotals.get(a)));

  const totalHours = round2_(sumMapValues_(dailyTotals));
  const dayCount = allDates.length;
  const averagePerDay = dayCount > 0 ? round2_(totalHours / dayCount) : 0;
  const completeDays = allDates.filter((d) => (dailyTotals.get(d) || 0) >= DAILY_TARGET_HOURS).length;
  const completionRate = dayCount > 0 ? round2_((completeDays / dayCount) * 100) : 0;

  const dash = getOrCreateDashboardSheet_(ss);
  dash.clear();
  dash.getCharts().forEach((chart) => dash.removeChart(chart));
  dash.setHiddenGridlines(true);

  // En-tête
  dash.getRange('A1:D1').merge().setValue('Dashboard Sandro - Monitoring Temps');
  dash.getRange('A2').setValue('Source');
  dash.getRange('B2').setValue(sourceSheet.getName());
  dash.getRange('C2').setValue('Dernière MAJ');
  dash.getRange('D2').setValue(new Date());

  dash.getRange('A1:D1').setFontWeight('bold').setFontSize(14).setBackground('#1F2937').setFontColor('#FFFFFF');
  dash.getRange('A2:D2').setFontWeight('bold').setBackground('#F3F4F6');
  dash.getRange('D2').setNumberFormat('dd/MM/yyyy HH:mm');

  // KPI
  dash.getRange('A4:D4').setValues([['Heures totales', 'Jours saisis', 'Moyenne / jour', 'Jours >= 8h']]);
  dash.getRange('A5:D5').setValues([[totalHours, dayCount, averagePerDay, `${completeDays} (${completionRate}%)`]]);
  dash.getRange('A4:D4').setFontWeight('bold').setBackground('#DBEAFE');
  dash.getRange('A5:D5').setFontSize(13).setFontWeight('bold');
  dash.getRange('A5:C5').setNumberFormat('0.00');

  // Tableau jour
  const dayHeaderRow = 8;
  const dayDataRow = dayHeaderRow + 1;
  const dayHeaders = [['Date', 'Heures', 'Objectif', 'Écart', 'Statut']];
  const dayRows = allDates.map((dateKey) => {
    const hours = round2_(dailyTotals.get(dateKey) || 0);
    const delta = round2_(hours - DAILY_TARGET_HOURS);
    const status = hours >= DAILY_TARGET_HOURS ? 'OK' : `Manque ${round2_(DAILY_TARGET_HOURS - hours)}h`;
    return [new Date(`${dateKey}T00:00:00`), hours, DAILY_TARGET_HOURS, delta, status];
  });
  dash.getRange(dayHeaderRow, 1, 1, 5).setValues(dayHeaders).setFontWeight('bold').setBackground('#E5E7EB');
  if (dayRows.length > 0) {
    dash.getRange(dayDataRow, 1, dayRows.length, 5).setValues(dayRows);
    dash.getRange(dayDataRow, 1, dayRows.length, 1).setNumberFormat('dd/MM/yyyy');
    dash.getRange(dayDataRow, 2, dayRows.length, 3).setNumberFormat('0.00');
  }

  // Tableau mois
  const monthHeaderRow = 8;
  const monthDataRow = monthHeaderRow + 1;
  const monthHeaders = [['Mois', 'Heures']];
  const monthRows = allMonths.map((m) => [m, round2_(monthlyTotals.get(m) || 0)]);
  dash.getRange(monthHeaderRow, 8, 1, 2).setValues(monthHeaders).setFontWeight('bold').setBackground('#E5E7EB');
  if (monthRows.length > 0) {
    dash.getRange(monthDataRow, 8, monthRows.length, 2).setValues(monthRows);
    dash.getRange(monthDataRow, 9, monthRows.length, 1).setNumberFormat('0.00');
  }

  // Tableau affaire
  const affairHeaderRow = 8;
  const affairDataRow = affairHeaderRow + 1;
  const affairHeaders = [['Affaire', 'Heures']];
  const affairRows = allAffairs.map((a) => [a, round2_(affairTotals.get(a) || 0)]);
  dash.getRange(affairHeaderRow, 11, 1, 2).setValues(affairHeaders).setFontWeight('bold').setBackground('#E5E7EB');
  if (affairRows.length > 0) {
    dash.getRange(affairDataRow, 11, affairRows.length, 2).setValues(affairRows);
    dash.getRange(affairDataRow, 12, affairRows.length, 1).setNumberFormat('0.00');
  }

  // Tableau empilé par affaire / jour
  const stackedHeaderRow = 8;
  const stackedDataRow = stackedHeaderRow + 1;
  const stackedHeaders = [['Date', ...allAffairs]];
  const stackedRows = allDates.map((dateKey) => {
    const byAffair = dailyByAffair.get(dateKey) || new Map();
    const values = allAffairs.map((a) => round2_(byAffair.get(a) || 0));
    return [new Date(`${dateKey}T00:00:00`), ...values];
  });
  const stackedCols = stackedHeaders[0].length;
  dash.getRange(stackedHeaderRow, 14, 1, stackedCols).setValues(stackedHeaders).setFontWeight('bold').setBackground('#E5E7EB');
  if (stackedRows.length > 0) {
    dash.getRange(stackedDataRow, 14, stackedRows.length, stackedCols).setValues(stackedRows);
    dash.getRange(stackedDataRow, 14, stackedRows.length, 1).setNumberFormat('dd/MM/yyyy');
    if (stackedCols > 1) dash.getRange(stackedDataRow, 15, stackedRows.length, stackedCols - 1).setNumberFormat('0.00');
  }

  // Mise en forme colonnes
  dash.setColumnWidth(1, 110);
  dash.setColumnWidth(2, 90);
  dash.setColumnWidth(3, 90);
  dash.setColumnWidth(4, 95);
  dash.setColumnWidth(5, 160);
  dash.setColumnWidth(8, 90);
  dash.setColumnWidth(9, 90);
  dash.setColumnWidth(11, 220);
  dash.setColumnWidth(12, 90);
  dash.setColumnWidth(14, 110);
  for (let c = 15; c < 15 + allAffairs.length; c++) {
    dash.setColumnWidth(c, 100);
  }

  // Conditionnelle écart
  if (dayRows.length > 0) {
    const deltaRange = dash.getRange(dayDataRow, 4, dayRows.length, 1);
    const rules = dash.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#FEE2E2')
        .setFontColor('#991B1B')
        .setRanges([deltaRange])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setBackground('#DCFCE7')
        .setFontColor('#166534')
        .setRanges([deltaRange])
        .build()
    );
    dash.setConditionalFormatRules(rules);
  }

  // Graphiques
  if (dayRows.length > 0) {
    // 1) Courbe Heures/Jour vs objectif
    dash.insertChart(
      dash
        .newChart()
        .asLineChart()
        .addRange(dash.getRange(dayHeaderRow, 1, dayRows.length + 1, 3))
        .setPosition(8, 6, 0, 0)
        .setOption('title', 'Heures par jour vs Objectif 8h')
        .setOption('legend', { position: 'bottom' })
        .setOption('curveType', 'function')
        .setOption('hAxis', { title: 'Date' })
        .setOption('vAxis', { title: 'Heures' })
        .build()
    );
  }

  if (monthRows.length > 0) {
    // 2) Histogramme heures par mois
    dash.insertChart(
      dash
        .newChart()
        .asColumnChart()
        .addRange(dash.getRange(monthHeaderRow, 8, monthRows.length + 1, 2))
        .setPosition(28, 6, 0, 0)
        .setOption('title', 'Heures par mois')
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Mois' })
        .setOption('vAxis', { title: 'Heures' })
        .build()
    );
  }

  if (affairRows.length > 0) {
    // 3) Répartition par affaire
    dash.insertChart(
      dash
        .newChart()
        .asPieChart()
        .addRange(dash.getRange(affairHeaderRow, 11, affairRows.length + 1, 2))
        .setPosition(8, 10, 0, 0)
        .setOption('title', 'Répartition des heures par affaire')
        .build()
    );
  }

  if (stackedRows.length > 0 && allAffairs.length > 0) {
    // 4) Vue empilée jour par affaire
    dash.insertChart(
      dash
        .newChart()
        .asColumnChart()
        .addRange(dash.getRange(stackedHeaderRow, 14, stackedRows.length + 1, stackedCols))
        .setPosition(28, 10, 0, 0)
        .setOption('title', 'Détail journalier par affaire (empilé)')
        .setOption('isStacked', true)
        .setOption('legend', { position: 'bottom' })
        .setOption('hAxis', { title: 'Date' })
        .setOption('vAxis', { title: 'Heures' })
        .build()
    );
  }

  SpreadsheetApp.flush();
  ss.toast(`Dashboard actualisé (${dayCount} jours, ${round2_(totalHours)}h).`, 'Sandro Dashboard', 6);
}

function getOrCreateDashboardSheet_(ss) {
  let sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DASHBOARD_SHEET_NAME);
  return sheet;
}

function findSourceSheet_(ss) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    if (sheet.getName() === DASHBOARD_SHEET_NAME) continue;
    const width = Math.max(1, sheet.getLastColumn());
    const header = sheet.getRange(1, 1, 1, width).getValues()[0].map((v) => String(v || '').trim());
    const hasDate = header.includes('Date');
    const hasAffair = header.includes('Affaire');
    const hasDuration = header.includes('Durée (h)');
    if (hasDate && hasAffair && hasDuration) return sheet;
  }
  return null;
}

function toDateKey_(value, tz, fallbackTimestamp) {
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  const direct = String(value || '').trim();
  const isoMatch = direct.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;

  const frMatch = direct.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (frMatch) return `${frMatch[3]}-${frMatch[2].padStart(2, '0')}-${frMatch[1].padStart(2, '0')}`;

  const fallback = String(fallbackTimestamp || '').trim();
  const fallbackFr = fallback.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fallbackFr) return `${fallbackFr[3]}-${fallbackFr[2].padStart(2, '0')}-${fallbackFr[1].padStart(2, '0')}`;

  return '';
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  const str = String(value || '').trim().replace(',', '.');
  if (!str) return NaN;
  return Number(str);
}

function incrementMap_(map, key, inc) {
  map.set(key, round2_((map.get(key) || 0) + inc));
}

function sumMapValues_(map) {
  let sum = 0;
  map.forEach((v) => {
    sum += Number(v) || 0;
  });
  return sum;
}

function round2_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}
