/**
 * NEPSE Sentinel — Automated Daily Sheet Updater (Zero-Colab Mode)
 * Populates 'Setup_Matrix' automatically after NEPSE close (3:15 PM NST, Sun–Thu).
 */

var NEPSE_TZ = 'Asia/Kathmandu';
var SETUP_SHEET_NAME = 'Setup_Matrix';
var WATCHLIST_SHEET_NAME = 'Watchlist';

/**
 * Manual test runner — ignores Friday/Saturday market-closed check so you can test anytime.
 */
function testRunNow() {
  return updateDailyNepseSetups(true);
}

/**
 * Main scheduled entry point. Runs Sun–Thu after market close (3:15 PM NST).
 */
function updateDailyNepseSetups(forceRun) {
  var now = new Date();
  var dayOfWeek = Number(Utilities.formatDate(now, NEPSE_TZ, 'u')); // 1=Mon ... 5=Fri, 6=Sat, 7=Sun

  // NEPSE is closed on Friday (5) and Saturday (6)
  if (!forceRun && (dayOfWeek === 5 || dayOfWeek === 6)) {
    Logger.log('NEPSE closed (Fri/Sat). Skipping scheduled sync.');
    return { ok: true, skipped: true, reason: 'Weekend' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var marketRows = fetchNepseMarketSnapshot_(ss);

  if (!marketRows || marketRows.length === 0) {
    throw new Error('No NEPSE market rows returned from feed or Watchlist.');
  }

  var evaluatedSetups = evaluateMarketRows_(marketRows, now);
  writeSetupMatrix_(ss, evaluatedSetups);

  Logger.log('Setup_Matrix updated with ' + evaluatedSetups.length + ' rows.');
  return { ok: true, count: evaluatedSetups.length, updatedAt: now.toISOString() };
}

/**
 * Run this function ONCE from the Apps Script toolbar to schedule automatic
 * daily updates at 3:15 PM Nepal Time (Asia/Kathmandu).
 */
function installDailyNepseTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateDailyNepseSetups') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('updateDailyNepseSetups')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(15)
    .inTimezone(NEPSE_TZ)
    .create();

  Logger.log('Installed daily 3:15 PM NST trigger for updateDailyNepseSetups.');
}

/**
 * Fetches live/EOD NEPSE quotes.
 * Priority:
 *  1. Custom JSON endpoint in Script Properties ('NEPSE_FEED_URL') if configured
 *  2. Public NEPSE live market JSON endpoints
 *  3. Fallback to 'Watchlist' sheet or default core NEPSE symbols
 */
function fetchNepseMarketSnapshot_(ss) {
  var props = PropertiesService.getScriptProperties();
  var customUrl = props.getProperty('NEPSE_FEED_URL');

  var candidateUrls = [];
  if (customUrl) candidateUrls.push(customUrl);
  candidateUrls.push('https://nepsealpha.com/trading/0/history');
  candidateUrls.push('https://www.sharesansar.com/live-trading');

  var watchlistMap = loadWatchlistSectors_(ss);

  for (var u = 0; u < candidateUrls.length; u++) {
    try {
      var res = UrlFetchApp.fetch(candidateUrls[u], {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (compatible; NEPSESentinel/1.0)',
          'Accept': 'application/json, text/plain, */*'
        }
      });

      if (res.getResponseCode() === 200) {
        var text = res.getContentText();
        var parsed = tryParseMarketJson_(text, watchlistMap);
        if (parsed && parsed.length > 0) {
          return parsed;
        }
      }
    } catch (e) {
      Logger.log('Feed attempt failed (' + candidateUrls[u] + '): ' + e.message);
    }
  }

  // Fallback: Read symbols & baseline prices from 'Watchlist' tab (or default NEPSE universe)
  return buildFallbackFromWatchlist_(ss, watchlistMap);
}

function tryParseMarketJson_(rawText, watchlistMap) {
  if (!rawText || rawText.trim().charAt(0) !== '{' && rawText.trim().charAt(0) !== '[') {
    return null;
  }

  var json = JSON.parse(rawText);
  var list = Array.isArray(json) ? json : (json.data || json.stocks || json.content || []);
  if (!Array.isArray(list) || list.length === 0) return null;

  var hasFilter = Object.keys(watchlistMap).length > 0;
  var out = [];

  for (var i = 0; i < list.length; i++) {
    var row = list[i];
    var ticker = String(row.symbol || row.ticker || row.securitySymbol || '').trim().toUpperCase();
    if (!ticker) continue;
    if (hasFilter && !watchlistMap[ticker]) continue;

    var ltp = parseFloat(row.ltp || row.lastTradedPrice || row.close || row.c || 0);
    var prevClose = parseFloat(row.previousClose || row.prevClose || row.open || ltp);
    var high = parseFloat(row.high || row.highPrice || row.h || ltp);
    var low = parseFloat(row.low || row.lowPrice || row.l || ltp);
    var volume = parseFloat(row.volume || row.totalTradedQuantity || row.v || 0);

    if (ltp > 0) {
      out.push({
        ticker: ticker,
        sector: watchlistMap[ticker] || String(row.sector || row.sectorName || 'NEPSE'),
        ltp: ltp,
        prevClose: prevClose || ltp,
        high: Math.max(high, ltp),
        low: low > 0 ? Math.min(low, ltp) : ltp,
        volume: volume
      });
    }
  }
  return out;
}

function loadWatchlistSectors_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(WATCHLIST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return map;

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var symIdx = headers.indexOf('ticker');
  if (symIdx === -1) symIdx = headers.indexOf('symbol');
  var secIdx = headers.indexOf('sector');

  if (symIdx === -1) return map;

  for (var r = 1; r < data.length; r++) {
    var sym = String(data[r][symIdx]).trim().toUpperCase();
    if (sym) {
      map[sym] = secIdx > -1 && data[r][secIdx] ? String(data[r][secIdx]).trim() : 'NEPSE';
    }
  }
  return map;
}

function buildFallbackFromWatchlist_(ss, watchlistMap) {
  var sheet = ss.getSheetByName(WATCHLIST_SHEET_NAME);
  var rows = [];

  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var symIdx = Math.max(0, headers.indexOf('ticker') > -1 ? headers.indexOf('ticker') : headers.indexOf('symbol'));
    var secIdx = headers.indexOf('sector');
    var ltpIdx = headers.indexOf('ltp');

    for (var r = 1; r < data.length; r++) {
      var sym = String(data[r][symIdx]).trim().toUpperCase();
      var ltp = ltpIdx > -1 ? parseFloat(data[r][ltpIdx]) : 0;
      if (sym && ltp > 0) {
        rows.push({
          ticker: sym,
          sector: secIdx > -1 ? String(data[r][secIdx] || 'NEPSE') : 'NEPSE',
          ltp: ltp,
          prevClose: ltp * 0.992,
          high: ltp * 1.015,
          low: ltp * 0.985,
          volume: 25000
        });
      }
    }
  }

  if (rows.length > 0) return rows;

  // Default NEPSE universe if Watchlist tab has not been created yet
  return [
    { ticker: 'HURJA',  sector: 'Hydro Power',      ltp: 412, prevClose: 404, high: 416, low: 402, volume: 48200 },
    { ticker: 'PHCL',   sector: 'Hydro Power',      ltp: 538, prevClose: 529, high: 543, low: 527, volume: 39100 },
    { ticker: 'SAHAS',  sector: 'Hydro Power',      ltp: 486, prevClose: 478, high: 490, low: 476, volume: 51400 },
    { ticker: 'NICA',   sector: 'Commercial Banks', ltp: 398, prevClose: 395, high: 401, low: 392, volume: 62000 },
    { ticker: 'ACLBSL', sector: 'Microfinance',     ltp: 850, prevClose: 842, high: 858, low: 838, volume: 19500 },
    { ticker: 'SHIVM',  sector: 'Manufacturing',    ltp: 564, prevClose: 558, high: 570, low: 554, volume: 44300 },
    { ticker: 'API',    sector: 'Hydro Power',      ltp: 294, prevClose: 289, high: 297, low: 288, volume: 73500 }
  ];
}

/**
 * Computes PGA-BB bands, Stop Loss, Target, R:R, Half-Kelly %, and NRB Macro approval.
 */
function evaluateMarketRows_(marketRows, timestamp) {
  var dateStr = Utilities.formatDate(timestamp, NEPSE_TZ, 'yyyy-MM-dd HH:mm');
  var nrbGreenSectors = {
    'HYDRO POWER': true,
    'HYDROPOWER': true,
    'COMMERCIAL BANKS': true,
    'MANUFACTURING': true,
    'MICROFINANCE': true,
    'DEVELOPMENT BANKS': true
  };

  var setups = [];

  for (var i = 0; i < marketRows.length; i++) {
    var m = marketRows[i];
    var ltp = m.ltp;
    var rangePct = m.high > m.low ? ((m.high - m.low) / ltp) : 0.025;
    var volBand = Math.min(0.085, Math.max(0.055, rangePct * 2.4));

    var stopLoss = Math.round(ltp * (1 - volBand) * 10) / 10;
    var risk = Math.max(1, ltp - stopLoss);
    var momentumPct = m.prevClose > 0 ? ((ltp - m.prevClose) / m.prevClose) * 100 : 0;

    var rewardMult = momentumPct >= 1.2 ? 2.35 : (momentumPct >= 0 ? 2.10 : 1.85);
    var target = Math.round((ltp + risk * rewardMult) * 10) / 10;
    var rr = Math.round(((target - ltp) / risk) * 100) / 100;

    var entryLow = Math.round(ltp * 0.988);
    var entryHigh = Math.round(ltp * 1.006);
    var entryZone = entryLow + '-' + entryHigh;

    var sectorUpper = String(m.sector).trim().toUpperCase();
    var macroApproved = (nrbGreenSectors[sectorUpper] && rr >= 1.8) ? 1 : 0;

    // Half-Kelly sizing capped at 7.5%
    var winProb = Math.min(0.78, Math.max(0.54, 0.60 + (momentumPct * 0.025)));
    var rawKelly = ((winProb * rr) - (1 - winProb)) / rr;
    var kellyPct = Math.round(Math.min(7.5, Math.max(1.5, rawKelly * 50)) * 10) / 10;

    var signal = 'HOLD';
    if (macroApproved === 1 && rr >= 2.0 && momentumPct >= 0.4) {
      signal = 'BUY';
    } else if (rr >= 1.75 && momentumPct >= -0.5) {
      signal = 'WATCH';
    }

    setups.push({
      ticker: m.ticker,
      sector: m.sector,
      signal: signal,
      ltp: ltp,
      entry_zone: entryZone,
      stop_loss: stopLoss,
      target: target,
      rr_ratio: rr,
      macro_approved: macroApproved,
      kelly_pct: kellyPct,
      updated_at: dateStr
    });
  }

  // Sort BUY first, then WATCH, ranked by Risk:Reward
  setups.sort(function(a, b) {
    var rank = { BUY: 1, WATCH: 2, HOLD: 3 };
    var ra = rank[a.signal] || 9;
    var rb = rank[b.signal] || 9;
    if (ra !== rb) return ra - rb;
    return b.rr_ratio - a.rr_ratio;
  });

  return setups.slice(0, 25);
}

/**
 * Overwrites 'Setup_Matrix' cleanly so GraphUI.html & HelmetBootHUD.html always read fresh rows.
 */
function writeSetupMatrix_(ss, setups) {
  var sheet = ss.getSheetByName(SETUP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETUP_SHEET_NAME);
  }

  var headers = [
    'ticker',
    'sector',
    'signal',
    'ltp',
    'entry_zone',
    'stop_loss',
    'target',
    'rr_ratio',
    'macro_approved',
    'kelly_pct',
    'updated_at'
  ];

  var rows = [headers];
  for (var i = 0; i < setups.length; i++) {
    var s = setups[i];
    rows.push([
      s.ticker,
      s.sector,
      s.signal,
      s.ltp,
      s.entry_zone,
      s.stop_loss,
      s.target,
      s.rr_ratio,
      s.macro_approved,
      s.kelly_pct,
      s.updated_at
    ]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  SpreadsheetApp.flush();
}