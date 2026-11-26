/**
 * NEPSE Sentinel Cloud — Dual-Brain Telemetry & Knowledge Graph Data Engine (Graph.gs)
 */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).toLowerCase() : '';

  if (action === 'status') {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'NEPSE Sentinel Dual-Brain Online',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'portfolio') {
    const graphData = buildLiveGraphData(true);
    const held = graphData.setups.filter(s => s.in_portfolio).map(s => s.ticker);
    return ContentService
      .createTextOutput(JSON.stringify({ held_tickers: held }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'telemetry') {
    const graphData = buildLiveGraphData(true);
    return ContentService
      .createTextOutput(JSON.stringify(graphData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  logTelemetry('doGet', 'WEB_HUD', 'OK', 'Opened Dual-Brain 3D Telemetry HUD');
  const graphData = buildLiveGraphData(false);
  const tpl = HtmlService.createTemplateFromFile('GraphUI');
  tpl.graphDataJson = JSON.stringify(graphData).replace(/</g, '\\u003c');

  return tpl.evaluate()
    .setTitle('NEPSE Sentinel — Dual-Brain 3D Telemetry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildLiveGraphData(skipLog) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('setup_matrix');
  const portfolioSheet = ss.getSheetByName('Portfolio');
  const l3Sheet = ss.getSheetByName('L3_Playbook_Config');
  const teleSheet = ss.getSheetByName('System_Telemetry');
  const l0Sheet = ss.getSheetByName('L0_Run_Snapshots');

  // 1. Portfolio Holdings
  const heldTickers = new Set();
  if (portfolioSheet && portfolioSheet.getLastRow() > 1) {
    const pData = portfolioSheet.getDataRange().getValues();
    const pHeaders = pData[0].map(h => String(h).trim().toLowerCase());
    const tIdx = Math.max(0, pHeaders.indexOf('ticker'));
    const qIdx = pHeaders.indexOf('qty') !== -1 ? pHeaders.indexOf('qty') : pHeaders.indexOf('quantity');
    for (let i = 1; i < pData.length; i++) {
      const sym = String(pData[i][tIdx] || '').trim().toUpperCase();
      const qty = qIdx !== -1 ? Number(pData[i][qIdx] || 0) : 1;
      if (sym && qty > 0) heldTickers.add(sym);
    }
  }

  // 2. Output Brain: NEPSE Setup Matrix
  const setups = [];
  if (matrixSheet && matrixSheet.getLastRow() > 1) {
    const rows = matrixSheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim().toLowerCase());
    const idx = (name) => headers.indexOf(name);

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const ticker = String(r[idx('ticker')] || '').trim().toUpperCase();
      if (!ticker) continue;
      setups.push({
        ticker: ticker,
        sector: String(r[idx('sector')] || 'OTHERS').trim().toUpperCase(),
        ltp: Number(r[idx('ltp')] || 0),
        signal: String(r[idx('signal')] || 'WATCH').trim().toUpperCase(),
        entry_zone: String(r[idx('entry_zone')] || ''),
        stop_loss: Number(r[idx('stop_loss')] || 0),
        target: Number(r[idx('target')] || 0),
        avg_turnover_20d: Number(r[idx('avg_turnover_20d')] || 0),
        macro_approved: Number(r[idx('macro_approved')] !== '' ? r[idx('macro_approved')] : 1),
        citation_chunk_id: String(r[idx('citation_chunk_id')] || '').trim(),
        bull_case: String(r[idx('bull_case')] || ''),
        bear_case: String(r[idx('bear_case')] || ''),
        risk_flag: String(r[idx('risk_flag')] || ''),
        kelly_pct: Number(r[idx('kelly_pct')] || 0),
        in_portfolio: heldTickers.has(ticker)
      });
    }
  }

  // 3. L3 Playbook Rules
  const playbooks = [];
  if (l3Sheet && l3Sheet.getLastRow() > 1) {
    const pbRows = l3Sheet.getDataRange().getValues();
    for (let i = 1; i < pbRows.length; i++) {
      if (String(pbRows[i][5] || '').toUpperCase() === 'ACTIVE') {
        playbooks.push({
          version: String(pbRows[i][0] || 'v1.0'),
          sector: String(pbRows[i][1] || 'ALL').toUpperCase(),
          rule: String(pbRows[i][4] || '')
        });
      }
    }
  }

  // 4. Input Brain: Live Function & Sub-Function Execution Telemetry
  const telemetryEvents = [];
  if (teleSheet && teleSheet.getLastRow() > 1) {
    const tRows = teleSheet.getDataRange().getValues();
    const start = Math.max(1, tRows.length - 120);
    for (let i = start; i < tRows.length; i++) {
      telemetryEvents.push({
        ts: String(tRows[i][0] || ''),
        nptTime: String(tRows[i][1] || ''),
        hourNpt: Number(tRows[i][2] !== '' ? tRows[i][2] : -1),
        fnName: String(tRows[i][3] || ''),
        parentFn: String(tRows[i][4] || ''),
        status: String(tRows[i][5] || 'OK'),
        detail: String(tRows[i][6] || ''),
        targetTicker: String(tRows[i][7] || '')
      });
    }
  }

  // Backfill historical L0_Run_Snapshots if telemetry has few entries yet
  if (telemetryEvents.length < 5 && l0Sheet && l0Sheet.getLastRow() > 1) {
    const l0Rows = l0Sheet.getDataRange().getValues();
    const start0 = Math.max(1, l0Rows.length - 25);
    for (let i = start0; i < l0Rows.length; i++) {
      const dt = l0Rows[i][0] instanceof Date ? l0Rows[i][0] : new Date(l0Rows[i][0]);
      if (isNaN(dt.getTime())) continue;
      const runType = String(l0Rows[i][1] || 'EOD_SYNC');
      const fn = runType === 'TELEGRAM_REPORT' ? 'buildTickerReport' : 'syncSetupMatrixFromColab';
      telemetryEvents.push({
        ts: dt.toISOString(),
        nptTime: Utilities.formatDate(dt, 'Asia/Kathmandu', 'yyyy-MM-dd HH:mm:ss'),
        hourNpt: Number(Utilities.formatDate(dt, 'Asia/Kathmandu', 'H')),
        fnName: fn,
        parentFn: 'doPost',
        status: 'OK',
        detail: `${runType} (${l0Rows[i][2]})`,
        targetTicker: String(l0Rows[i][2] || '')
      });
    }
  }

  const now = new Date();
  return {
    updatedAt: Utilities.formatDate(now, 'Asia/Kathmandu', 'yyyy-MM-dd HH:mm:ss') + ' NPT',
    currentHourNpt: Number(Utilities.formatDate(now, 'Asia/Kathmandu', 'H')),
    setups: setups,
    playbooks: playbooks,
    telemetry: telemetryEvents
  };
}