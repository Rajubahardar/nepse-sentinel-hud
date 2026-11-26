/**
 * NEPSE Sentinel — Mark VII J.A.R.V.I.S. Peripheral Helmet Visor Service
 * Integrates with GraphUI.html via <?!= include('HelmetBootHUD'); ?>
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Hydrates overall Trade_Journal & Setup_Matrix backtest metrics into HelmetBootHUD.html.
 */
function getHelmetBacktestJson() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var journalSheet = ss.getSheetByName('Trade_Journal');
  var setupSheet = ss.getSheetByName('Setup_Matrix');

  var wins = 0;
  var totalClosed = 0;

  if (journalSheet && journalSheet.getLastRow() > 1) {
    var jData = journalSheet.getDataRange().getValues();
    var jHeaders = jData[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var pnlIdx = jHeaders.indexOf('pnl_pct');
    if (pnlIdx === -1) pnlIdx = jHeaders.indexOf('return_pct');
    var statusIdx = jHeaders.indexOf('status');

    for (var i = 1; i < jData.length; i++) {
      var row = jData[i];
      var status = statusIdx > -1 ? String(row[statusIdx]).trim().toUpperCase() : 'CLOSED';
      var pnl = pnlIdx > -1 ? parseFloat(row[pnlIdx]) : NaN;
      if (!isNaN(pnl) && (status === 'CLOSED' || status === 'WIN' || status === 'LOSS' || statusIdx === -1)) {
        totalClosed++;
        if (pnl > 0) wins++;
      }
    }
  }

  var setupCount = (setupSheet && setupSheet.getLastRow() > 1) ? (setupSheet.getLastRow() - 1) : 0;
  var accuracy = totalClosed > 0 ? Math.round((wins / totalClosed) * 1000) / 10 : null;

  return JSON.stringify({
    ok: true,
    journalClosed: totalClosed,
    journalWins: wins,
    journalAccuracy: accuracy,
    setupCount: setupCount,
    timestamp: new Date().toISOString()
  });
}