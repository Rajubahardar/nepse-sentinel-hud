/**
- NEPSE Sentinel Cloud — Core Backend, WACC Engine, L2 Guards, Telegram & Live Telemetry
- Paired with: Graph.gs + GraphUI.html + VoiceCommandRouter.html (Dual-Brain 3D Telemetry HUD)
 */

const PROPS = PropertiesService.getScriptProperties();

// ==========================================
// 1. ONE-TIME SHEET INITIALIZER (RUN ONCE)
// ==========================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = {
    'Portfolio': ['ticker', 'qty', 'buy_price', 'true_wacc', 'buy_date', 'settlement_status', 'sector'],
    'setup_matrix': [
      'ticker', 'sector', 'ltp', 'signal', 'entry_zone', 'stop_loss', 'target',
      'avg_turnover_20d', 'is_price_adjusted_today', 'macro_approved',
      'citation_chunk_id', 'bull_case', 'bear_case', 'risk_flag', 'kelly_pct', 'updated_at'
    ],
    'L0_Run_Snapshots': ['timestamp', 'run_type', 'ticker', 'playbook_version', 'input_json', 'guard_result_json'],
    'L1_Scoped_Lessons': ['lesson_id', 'scope_sector', 'failure_pattern', 'rule_note', 'active'],
    'L3_Playbook_Config': ['version', 'sector', 'min_turnover_20d', 'max_kelly_pct', 'playbook_prompt', 'pareto_status'],
    'Pareto_Archive': ['config_version', 'net_return_pct', 'max_drawdown_pct', 'compliance_pass_rate', 'promoted'],
    'System_Telemetry': ['timestamp', 'npt_time', 'hour_npt', 'fn_name', 'parent_fn', 'status', 'detail', 'target_ticker']
  };

  Object.entries(schemas).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  });

  const l3 = ss.getSheetByName('L3_Playbook_Config');
  if (l3.getLastRow() === 1) {
    l3.appendRow(['v1.0', 'ALL', 5000000, 0.15, 'Default NEPSE swing & fundamental rules', 'ACTIVE']);
    l3.appendRow(['v1.0', 'HYDRO', 8000000, 0.10, 'Check promoter lock-in expiry & hydrology season', 'ACTIVE']);
    l3.appendRow(['v1.0', 'BANKING', 5000000, 0.20, 'Verify NRB CD ratio, NPL trend & dividend cap', 'ACTIVE']);
  }

  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  logTelemetry('setupSheets', 'SYS_INIT', 'OK', 'Initialized/verified all 7 sheets');
}

// ==========================================
// 1B. REAL-TIME FUNCTION TELEMETRY LOGGER
// ==========================================
function logTelemetry(fnName, parentFn, status, detail, targetTicker) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('System_Telemetry');
    if (!sheet) {
      sheet = ss.insertSheet('System_Telemetry');
      sheet.appendRow(['timestamp', 'npt_time', 'hour_npt', 'fn_name', 'parent_fn', 'status', 'detail', 'target_ticker']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
    const now = new Date();
    const nptStr = Utilities.formatDate(now, 'Asia/Kathmandu', 'yyyy-MM-dd HH:mm:ss');
    const hourNpt = Number(Utilities.formatDate(now, 'Asia/Kathmandu', 'H'));
    sheet.appendRow([
      now.toISOString(),
      nptStr,
      hourNpt,
      fnName,
      parentFn || 'ROOT_INPUT',
      status || 'OK',
      String(detail || '').slice(0, 180),
      String(targetTicker || '').toUpperCase()
    ]);

    // Keep telemetry sheet lightweight (rolling last 200 runs)
    const lastRow = sheet.getLastRow();
    if (lastRow > 240) {
      sheet.deleteRows(2, lastRow - 201);
    }
  } catch (e) {
    console.warn('Telemetry skip:', e);
  }
}

// ==========================================
// 2. EXACT 2026 NEPSE FEE, WACC & CGT ENGINE
// ==========================================
function getNepseCommission(amount) {
  let rate = 0.0024;
  if (amount <= 50000) rate = 0.0036;
  else if (amount <= 500000) rate = 0.0033;
  else if (amount <= 2000000) rate = 0.0031;
  else if (amount <= 10000000) rate = 0.0027;
  return Math.max(10, amount * rate);
}

function calcNepseBuyWacc(qty, price) {
  const tradeVal = qty * price;
  const comm = getNepseCommission(tradeVal);
  const sebon = tradeVal * 0.00015;
  const dp = 25.0;
  const totalCost = tradeVal + comm + sebon + dp;
  const res = {
    tradeVal,
    totalCost: Number(totalCost.toFixed(2)),
    trueWacc: Number((totalCost / qty).toFixed(2))
  };
  logTelemetry('calcNepseBuyWacc', 'logPaperOrRealBuy', 'OK', `Qty ${qty} @ Rs ${price} -> WACC Rs ${res.trueWacc}`);
  return res;
}

function calcNepseNetSell(qty, sellPrice, trueWacc, heldDays) {
  const saleVal = qty * sellPrice;
  const comm = getNepseCommission(saleVal);
  const sebon = saleVal * 0.00015;
  const dp = 25.0;
  const netBeforeTax = saleVal - comm - sebon - dp;
  const costBasis = qty * trueWacc;
  const gain = netBeforeTax - costBasis;
  // Individual CGT rates effective Sept 22, 2026: 3.75% (>365d), 5% (<=365d)
  const cgtRate = heldDays > 365 ? 0.0375 : 0.05;
  const cgt = gain > 0 ? gain * cgtRate : 0;
  const res = {
    netReceivable: Number((netBeforeTax - cgt).toFixed(2)),
    netProfit: Number((gain - cgt).toFixed(2)),
    cgtPaid: Number(cgt.toFixed(2))
  };
  logTelemetry('calcNepseNetSell', 'ENGINE_FIN', 'OK', `Net Rs ${res.netReceivable} (CGT Rs ${res.cgtPaid})`);
  return res;
}

// ==========================================
// 3. REELS-RSI L2 DETERMINISTIC GUARDS & TESTS
// ==========================================
function runL2Guards(row) {
  const failures = [];
  if (Number(row.avg_turnover_20d) < 5000000) {
    failures.push('ILLIQUID_COUNTER (<50L turnover)');
  }
  if (Number(row.macro_approved) === 0 && !String(row.citation_chunk_id || '').trim()) {
    failures.push('UNGROUNDED_MACRO_BLOCK');
  }
  if (String(row.is_price_adjusted_today).toUpperCase() === 'TRUE' && String(row.signal).toUpperCase() === 'BREAKOUT') {
    failures.push('ADJUSTED_PRICE_FALSE_BREAKOUT');
  }
  if (Number(row.kelly_pct) > 0.25) {
    failures.push('KELLY_CAP_EXCEEDED (>25%)');
  }
  return { passed: failures.length === 0, failures };
}

function testL2Guards() {
  const mustPass = runL2Guards({
    avg_turnover_20d: 12000000,
    macro_approved: 1,
    citation_chunk_id: '',
    is_price_adjusted_today: false,
    signal: 'BREAKOUT',
    kelly_pct: 0.12
  });
  const mustFail = runL2Guards({
    avg_turnover_20d: 800000,
    macro_approved: 0,
    citation_chunk_id: '',
    is_price_adjusted_today: true,
    signal: 'BREAKOUT',
    kelly_pct: 0.30
  });
  if (!mustPass.passed || mustFail.failures.length !== 4) {
    logTelemetry('testL2Guards', 'SYS_INIT', 'ERR', 'Unit tests failed');
    throw new Error('L2 Unit Tests FAILED: ' + JSON.stringify({ mustPass, mustFail }));
  }
  logTelemetry('testL2Guards', 'SYS_INIT', 'OK', 'All L2 deterministic guard unit tests passed');
  return 'ALL L2 UNIT TESTS PASSED';
}

function runRetroactiveL2Audit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('L0_Run_Snapshots');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    logTelemetry('runRetroactiveL2Audit', 'doPost', 'OK', 'L0_Run_Snapshots empty');
    return 'L0_Run_Snapshots is empty — nothing to audit yet.';
  }

  let defects = 0;
  const details = [];
  for (let i = 1; i < data.length; i++) {
    const snap = JSON.parse(data[i][4] || '{}');
    const res = runL2Guards(snap);
    if (!res.passed) {
      defects++;
      details.push(`Row ${i + 1} (${data[i][2]}): ${res.failures.join(', ')}`);
    }
  }
  const msg = defects === 0
    ? `Retroactive Audit Clean: 0 defects across ${data.length - 1} historical runs.`
    : `Caught ${defects} historical defect(s):\n` + details.slice(0, 10).join('\n');
  logTelemetry('runRetroactiveL2Audit', 'doPost', defects === 0 ? 'OK' : 'WARN', msg);
  return msg;
}

// ==========================================
// 4. WEBHOOK (TELEGRAM + COLAB EOD SYNC)
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData) return HtmlService.createHtmlOutput('OK');
    const body = JSON.parse(e.postData.contents);

    // Route A: Colab Post-Market EOD Sync
    if (body.action === 'sync_eod' && body.secret === PROPS.getProperty('SYNC_SECRET')) {
      logTelemetry('doPost', 'WEBHOOK_COLAB', 'OK', `Received sync_eod (${(body.rows || []).length} rows)`);
      syncSetupMatrixFromColab(body.rows || []);
      return HtmlService.createHtmlOutput('SYNCED');
    }

    // Route B: Telegram Bot Update
    const cache = CacheService.getScriptCache();
    if (body.update_id && cache.get('tg_' + body.update_id)) {
      return HtmlService.createHtmlOutput('OK');
    }
    if (body.update_id) cache.put('tg_' + body.update_id, '1', 300);

    const msg = body.message;
    if (!msg || !msg.text) return HtmlService.createHtmlOutput('OK');

    const chatId = String(msg.chat.id);
    if (chatId !== PROPS.getProperty('ALLOWED_CHAT_ID')) {
      return HtmlService.createHtmlOutput('OK');
    }

    const parts = msg.text.trim().split(/\s+/);
    const cmd = parts[0].toLowerCase();
    logTelemetry('doPost', 'WEBHOOK_TG', 'OK', `Telegram command: ${msg.text.trim()}`, parts[1] || '');

    if (cmd === '/report' && parts[1]) {
      sendTelegram(chatId, buildTickerReport(parts[1].toUpperCase()), true, 'buildTickerReport');
    } else if (cmd === '/matrix') {
      sendTelegram(chatId, buildMatrixDigest('doPost'), true, 'buildMatrixDigest');
    } else if (cmd === '/graph') {
      const webUrl = PROPS.getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl();
      sendTelegram(chatId, 'NEPSE Sentinel Dual-Brain Telemetry:\n' + webUrl, false, 'doPost');
    } else if (cmd === '/buy' && parts.length >= 4) {
      sendTelegram(chatId, logPaperOrRealBuy(parts[1].toUpperCase(), Number(parts[2]), Number(parts[3]), parts[4] || 'GENERAL'), true, 'logPaperOrRealBuy');
    } else if (cmd === '/edis') {
      sendTelegram(chatId, checkEdisObligations('doPost'), true, 'checkEdisObligations');
    } else if (cmd === '/audit') {
      sendTelegram(chatId, runRetroactiveL2Audit(), true, 'runRetroactiveL2Audit');
    } else {
      sendTelegram(chatId, [
        '*NEPSE Sentinel Commands:*',
        '`/report NABIL` — Bull/Bear/Risk + L2 guard report',
        '`/matrix` — Today\'s Pareto-approved setups',
        '`/graph` — Open Dual-Brain 3D Telemetry HUD',
        '`/buy NABIL 100 542 BANKING` — Log trade & exact WACC',
        '`/edis` — Check T+0 / T+1 EDIS settlement locks',
        '`/audit` — Run retroactive L2 audit on L0 history'
      ].join('\n'), true, 'doPost');
    }
  } catch (err) {
    console.error(err);
    logTelemetry('doPost', 'WEBHOOK_ERR', 'ERR', err.message || String(err));
  }
  return HtmlService.createHtmlOutput('OK');
}

function sendTelegram(chatId, text, useMarkdown = true, callerFn = 'doPost') {
  const token = PROPS.getProperty('BOT_TOKEN');
  const url = `https://api.telegram.org/bot${token}/sendMessage`;

  if (useMarkdown) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'Markdown' }),
      muteHttpExceptions: true
    });
    const body = JSON.parse(res.getContentText() || '{}');
    if (body.ok) {
      logTelemetry('sendTelegram', callerFn, 'OK', `Sent ${text.length} chars (Markdown)`);
      return;
    }
  }

  const cleanText = useMarkdown ? text.replace(/[*_`]/g, '') : text;
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: cleanText }),
    muteHttpExceptions: true
  });
  logTelemetry('sendTelegram', callerFn, 'OK', `Sent ${cleanText.length} chars (Plain)`);
}

// ==========================================
// 5. SHEET & DIGEST HELPERS
// ==========================================
function logPaperOrRealBuy(ticker, qty, price, sector) {
  const calc = calcNepseBuyWacc(qty, price);
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio');
    sheet.appendRow([ticker, qty, price, calc.trueWacc, new Date().toISOString().slice(0, 10), 'LOCKED_T0', sector]);
  } finally {
    lock.releaseLock();
  }
  logTelemetry('logPaperOrRealBuy', 'doPost', 'OK', `Logged ${ticker} (${qty} @ ${price})`, ticker);
  return `Logged *${ticker}*: ${qty} @ Rs ${price}\nTotal Cost: Rs ${calc.totalCost} | *True WACC: Rs ${calc.trueWacc}* (Status: LOCKED_T0)`;
}

function buildTickerReport(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName('setup_matrix').getDataRange().getValues();
  const headers = data[0];
  const rowArr = data.slice(1).find(r => String(r[0]).toUpperCase() === ticker);
  if (!rowArr) {
    logTelemetry('buildTickerReport', 'doPost', 'WARN', `Ticker ${ticker} not in setup_matrix`, ticker);
    return `No entry for *${ticker}* in setup_matrix yet.`;
  }

  const row = Object.fromEntries(headers.map((h, i) => [h, rowArr[i]]));
  const guard = runL2Guards(row);

  ss.getSheetByName('L0_Run_Snapshots').appendRow([
    new Date(), 'TELEGRAM_REPORT', ticker, 'v1.0', JSON.stringify(row), JSON.stringify(guard)
  ]);
  logTelemetry('buildTickerReport', 'doPost', guard.passed ? 'OK' : 'WARN', `${ticker}: ${row.signal} (L2=${guard.passed})`, ticker);

  return [
    `*${ticker}* (${row.sector}) — LTP: Rs ${row.ltp}`,
    `Signal: *${row.signal}* | Entry: ${row.entry_zone} | SL: ${row.stop_loss} | Target: ${row.target}`,
    `Bull: ${row.bull_case || 'N/A'}`,
    `Bear: ${row.bear_case || 'N/A'}`,
    `Risk: ${row.risk_flag || 'None'} | Kelly: ${(Number(row.kelly_pct || 0) * 100).toFixed(1)}%`,
    `Macro Approved: ${row.macro_approved} (Chunk: ${row.citation_chunk_id || 'N/A'})`,
    `L2 Guard: *${guard.passed ? 'PASS' : 'BLOCKED -> ' + guard.failures.join(', ')}*`
  ].join('\n');
}

function buildMatrixDigest(callerFn = 'doPost') {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setup_matrix').getDataRange().getValues();
  if (data.length <= 1) return 'setup_matrix is empty. Run /eod sync from Colab after 3:05 PM.';
  const headers = data[0];
  const lines = ['*Approved NEPSE Setups (L2 + Macro Verified):*'];

  for (let i = 1; i < data.length; i++) {
    const row = Object.fromEntries(headers.map((h, idx) => [h, data[i][idx]]));
    if (String(row.signal).toUpperCase() === 'SCAN') continue;
    const guard = runL2Guards(row);
    if (Number(row.macro_approved) === 1 && guard.passed) {
      lines.push(`• *${row.ticker}* (${row.signal}) @ ${row.entry_zone} | SL ${row.stop_loss} | Kelly ${(Number(row.kelly_pct) * 100).toFixed(0)}%`);
    }
  }
  logTelemetry('buildMatrixDigest', callerFn, 'OK', `Compiled ${lines.length - 1} approved setups`);
  return lines.length > 1 ? lines.join('\n') : 'No setups passed both Macro Gate and L2 Guards today.';
}

function checkEdisObligations(callerFn = 'doPost') {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio').getDataRange().getValues();
  const pending = data.slice(1).filter(r => String(r[5]).includes('LOCKED') || String(r[5]).includes('EDIS'));
  logTelemetry('checkEdisObligations', callerFn, 'OK', `${pending.length} pending settlement lock(s)`);
  if (pending.length === 0) return 'No pending T+0 / T+1 EDIS settlement locks.';
  return '*Settlement Tracker:*\n' + pending.map(r => `• *${r[0]}* (${r[1]} units @ WACC ${r[3]}) — ${r[5]}`).join('\n');
}

function syncSetupMatrixFromColab(rows) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('setup_matrix');
    const l0 = ss.getSheetByName('L0_Run_Snapshots');
    const now = new Date();

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    if (!rows || rows.length === 0) return;

    const matrixBatch = [];
    const l0Batch = [];
    let debatedCount = 0;

    rows.forEach(r => {
      matrixBatch.push([
        r.ticker, r.sector, r.ltp, r.signal, r.entry_zone, r.stop_loss, r.target,
        r.avg_turnover_20d, r.is_price_adjusted_today, r.macro_approved,
        r.citation_chunk_id, r.bull_case, r.bear_case, r.risk_flag, r.kelly_pct, now
      ]);
      if (String(r.signal).toUpperCase() !== 'SCAN') {
        debatedCount++;
        const guard = runL2Guards(r);
        l0Batch.push([now, 'EOD_SYNC', r.ticker, 'v1.0', JSON.stringify(r), JSON.stringify(guard)]);
      }
    });

    sheet.getRange(2, 1, matrixBatch.length, matrixBatch[0].length).setValues(matrixBatch);
    if (l0Batch.length > 0) {
      l0.getRange(l0.getLastRow() + 1, 1, l0Batch.length, l0Batch[0].length).setValues(l0Batch);
    }
    logTelemetry('runL2Guards', 'syncSetupMatrixFromColab', 'OK', `Validated ${debatedCount} AI-debated setups`);
    logTelemetry('syncSetupMatrixFromColab', 'doPost', 'OK', `Synced ${matrixBatch.length} NEPSE securities to Output Cortex`);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 6. DEPLOYMENT & TRIGGER HELPERS
// ==========================================
function registerTelegramWebhook() {
  const token = PROPS.getProperty('BOT_TOKEN');
  const url = PROPS.getProperty('WEB_APP_URL');
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(url)}`);
  logTelemetry('registerTelegramWebhook', 'SYS_INIT', 'OK', 'Registered Telegram webhook');
  console.log(res.getContentText());
}

function setupDailyTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // 10:00 AM NPT Settlement & EDIS Reminder
  ScriptApp.newTrigger('pushMorningAlert').timeBased().everyDays(1).atHour(10).nearMinute(0).create();
  // 3:30 PM NPT Post-Market Setup Matrix Digest
  ScriptApp.newTrigger('pushEodDigest').timeBased().everyDays(1).atHour(15).nearMinute(30).create();
  logTelemetry('setupDailyTriggers', 'SYS_INIT', 'OK', 'Scheduled 10:00 AM & 3:30 PM NPT triggers');
}

function pushMorningAlert() {
  logTelemetry('pushMorningAlert', 'TRIG_10AM', 'OK', '10:00 AM NPT Scheduled Trigger fired');
  sendTelegram(PROPS.getProperty('ALLOWED_CHAT_ID'), checkEdisObligations('pushMorningAlert'), true, 'pushMorningAlert');
}

function pushEodDigest() {
  logTelemetry('pushEodDigest', 'TRIG_330PM', 'OK', '3:30 PM NPT Scheduled Trigger fired');
  sendTelegram(PROPS.getProperty('ALLOWED_CHAT_ID'), buildMatrixDigest('pushEodDigest'), true, 'pushEodDigest');
}

function checkWebhookStatus() {
  const token = PROPS.getProperty('BOT_TOKEN');
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/getWebhookInfo`);
  console.log(res.getContentText());
}

function debugTelegramBot() {
  const all = PROPS.getProperties();
  Object.entries(all).forEach(([k, v]) => PROPS.setProperty(k, String(v).trim()));

  const token = PROPS.getProperty('BOT_TOKEN');
  const chatId = PROPS.getProperty('ALLOWED_CHAT_ID');
  const webAppUrl = PROPS.getProperty('WEB_APP_URL');

  const info = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/getWebhookInfo`, { muteHttpExceptions: true });
  console.log('Webhook Info:', info.getContentText());

  if (webAppUrl && webAppUrl.endsWith('/exec')) {
    const setRes = UrlFetchApp.fetch(
      `https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(webAppUrl)}&drop_pending_updates=true`,
      { muteHttpExceptions: true }
    );
    console.log('Set Webhook:', setRes.getContentText());
  }

  sendTelegram(chatId, 'Diagnostic check: Dual-Brain Telemetry & Bot connection online!', false, 'debugTelegramBot');
  logTelemetry('debugTelegramBot', 'SYS_INIT', 'OK', 'Ran end-to-end bot diagnostic');
}

// ==========================================
// 7. VOICE HUD TELEMETRY & LIVE QUERY BRIDGE
// ==========================================
function logVoiceCommandTelemetry(commandText, actionTaken, targetTicker) {
  logTelemetry(
    'voiceCommandRouter',
    'JARVIS_VOICE_HUD',
    'OK',
    `${actionTaken || 'EXEC'}: ${String(commandText || '').trim()}`,
    targetTicker || ''
  );
  return true;
}

function getVoiceHudTopLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('setup_matrix');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const leads = [];

  for (let i = 1; i < data.length; i++) {
    const row = Object.fromEntries(headers.map((h, idx) => [h, data[i][idx]]));
    const sig = String(row.signal || '').toUpperCase();
    if (sig === 'SCAN') continue;

    const guard = runL2Guards(row);
    const ltp = Number(row.ltp || 0);
    const sl = Number(row.stop_loss || 0);
    const tgt = Number(row.target || 0);
    const risk = Math.abs(ltp - sl);
    const reward = Math.abs(tgt - ltp);
    const rr = risk > 0 ? Number((reward / risk).toFixed(2)) : 0;

    leads.push({
      ticker: String(row.ticker || '').toUpperCase(),
      sector: row.sector || '',
      signal: sig,
      ltp: ltp,
      stop_loss: sl,
      target: tgt,
      kelly_pct: Number(row.kelly_pct || 0),
      rr_ratio: rr,
      l2_passed: guard.passed,
      macro_approved: Number(row.macro_approved || 0) === 1
    });
  }

  leads.sort((a, b) => {
    const aBuy = a.signal === 'BUY' && a.l2_passed ? 1 : 0;
    const bBuy = b.signal === 'BUY' && b.l2_passed ? 1 : 0;
    if (bBuy !== aBuy) return bBuy - aBuy;
    return b.rr_ratio - a.rr_ratio;
  });

  logTelemetry('getVoiceHudTopLeads', 'JARVIS_VOICE_HUD', 'OK', `Returned ${leads.length} ranked setups`);
  return leads;
}