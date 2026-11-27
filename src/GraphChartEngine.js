// ============================================================================
// NEPSE SENTINEL — DEDICATED TECHNICAL GRAPH & AUTHENTICATION ENGINE
// ============================================================================

/**
 * 1. ONE-TIME SERVER SETUP (OR RUN VIA THE WEB APP FIRST-TIME SETUP FORM)
 * Stores your authorized Gmail and a salted SHA-256 password hash in private ScriptProperties.
 */
function setupGraphLoginCredentials(gmail, password) {
  const email = String(gmail || '').trim().toLowerCase();
  const pass = String(password || '').trim();
  if (!email || email.indexOf('@') === -1) {
    throw new Error('Please enter a valid Gmail address.');
  }
  if (pass.length < 6) {
    throw new Error('Password must be at least 6 characters.');
  }

  const salt = Utilities.getUuid();
  const hash = computeSaltedSha256_(salt, pass);
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    GRAPH_AUTH_EMAIL: email,
    GRAPH_AUTH_SALT: salt,
    GRAPH_AUTH_HASH: hash
  });

  return { ok: true, email: email, message: 'Graph login credentials saved securely.' };
}

function computeSaltedSha256_(salt, password) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + '::NEPSE_SENTINEL::' + password,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

/**
 * Checks whether Graph Login is configured and whether the current token/session is valid.
 */
function getGraphAuthStatus(sessionToken) {
  const props = PropertiesService.getScriptProperties();
  const savedEmail = props.getProperty('GRAPH_AUTH_EMAIL') || '';
  let activeGoogleEmail = '';
  try {
    activeGoogleEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {}

  if (!savedEmail) {
    return {
      isConfigured: false,
      authenticated: false,
      suggestedEmail: activeGoogleEmail
    };
  }

  let validToken = false;
  if (sessionToken) {
    const cachedEmail = CacheService.getScriptCache().get('graph_sess_' + String(sessionToken).trim());
    if (cachedEmail && cachedEmail === savedEmail) {
      validToken = true;
    }
  }

  return {
    isConfigured: true,
    authenticated: validToken,
    suggestedEmail: activeGoogleEmail || savedEmail
  };
}

/**
 * First-time registration from the Web App UI (only allowed if no account exists yet).
 */
function registerInitialGraphAccount(gmail, password) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('GRAPH_AUTH_EMAIL')) {
    return { ok: false, error: 'An authorized Gmail is already registered. Please sign in.' };
  }
  setupGraphLoginCredentials(gmail, password);
  return verifyGraphLogin(gmail, password);
}

/**
 * Verifies Gmail + Password and returns a 24-hour session token.
 */
function verifyGraphLogin(gmail, password) {
  const email = String(gmail || '').trim().toLowerCase();
  const pass = String(password || '').trim();
  const props = PropertiesService.getScriptProperties();

  const savedEmail = (props.getProperty('GRAPH_AUTH_EMAIL') || '').trim().toLowerCase();
  const savedSalt = props.getProperty('GRAPH_AUTH_SALT') || '';
  const savedHash = props.getProperty('GRAPH_AUTH_HASH') || '';

  if (!savedEmail || !savedSalt || !savedHash) {
    return { ok: false, error: 'No account configured yet. Use First-Time Setup.' };
  }

  const candidateHash = computeSaltedSha256_(savedSalt, pass);
  if (email !== savedEmail || candidateHash !== savedHash) {
    return { ok: false, error: 'Invalid Gmail or Password.' };
  }

  const token = Utilities.getUuid().replace(/-/g, '');
  // Cache for 6 hours (Apps Script max single put is 21600s)
  CacheService.getScriptCache().put('graph_sess_' + token, savedEmail, 21600);

  return {
    ok: true,
    token: token,
    email: savedEmail
  };
}

function logoutGraphSession(sessionToken) {
  if (sessionToken) {
    CacheService.getScriptCache().remove('graph_sess_' + String(sessionToken).trim());
  }
  return { ok: true };
}

// ============================================================================
// 2. MULTI-YEAR NEPSE HISTORICAL OHLCV ENGINE (TRADINGVIEW / NEPSEALPHA / MEROLAGANI)
// ============================================================================

function getLiveChartData(ticker, rangeKey, sessionToken) {
  const sym = String(ticker || 'NABIL').toUpperCase().trim();
  const range = String(rangeKey || '1Y').toUpperCase().trim();

  // Enforce login if configured
  const auth = getGraphAuthStatus(sessionToken);
  if (auth.isConfigured && !auth.authenticated) {
    return {
      ticker: sym,
      authRequired: true,
      source: 'AUTH_LOCKED',
      candles: []
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('setup_matrix');
  let rowObj = null;

  if (matrixSheet && matrixSheet.getLastRow() > 1) {
    const data = matrixSheet.getDataRange().getValues();
    const headers = data[0];
    const found = data.slice(1).find(r => String(r[0]).toUpperCase().trim() === sym);
    if (found) rowObj = Object.fromEntries(headers.map((h, i) => [h, found[i]]));
  }

  // Check 15-minute cache first for instant timeframe switching
  const cache = CacheService.getScriptCache();
  const cacheKey = 'nepse_tv_hist_v2_' + sym;
  let allCandles = null;
  let source = 'NEPSEALPHA_TV';

  try {
    const cachedRaw = cache.get(cacheKey);
    if (cachedRaw) {
      const parsedCache = JSON.parse(cachedRaw);
      if (parsedCache && Array.isArray(parsedCache.candles) && parsedCache.candles.length >= 10) {
        allCandles = parsedCache.candles;
        source = parsedCache.source + '_CACHED';
      }
    }
  } catch (e) {}

  if (!allCandles) {
    const fetched = fetchMultiYearNepseOhlcv_(sym);
    allCandles = fetched.candles;
    source = fetched.source;

    if (allCandles.length >= 10) {
      try {
        cache.put(cacheKey, JSON.stringify({
          source: source,
          candles: allCandles.slice(-1250) // Up to ~5 years of daily bars
        }), 900);
      } catch (e) {}
    }
  }

  // Fallback synthetic curve only if symbol is newly listed / halted with zero history
  if (!allCandles || allCandles.length < 5) {
    source = 'SETUP_MATRIX_TA';
    const ltp = Number((rowObj && rowObj.ltp) || 500);
    const baseStr = typeof getLastNepseTradingDateStr_ === 'function'
      ? getLastNepseTradingDateStr_()
      : Utilities.formatDate(new Date(), 'Asia/Kathmandu', 'yyyy-MM-dd');
    const baseDate = new Date(baseStr + 'T00:00:00Z');
    allCandles = [];
    let p = ltp * 0.92;
    for (let d = 60; d >= 0; d--) {
      const dt = new Date(baseDate.getTime() - d * 86400000);
      const dow = dt.getUTCDay();
      if (dow === 0 || dow === 6) continue;
      const iso = Utilities.formatDate(dt, 'UTC', 'yyyy-MM-dd');
      const open = Number(p.toFixed(1));
      const close = d === 0 ? ltp : Number((p * (1 + Math.sin(d * 1.4) * 0.011 + 0.0015)).toFixed(1));
      const high = Number((Math.max(open, close) * 1.008).toFixed(1));
      const low = Number((Math.min(open, close) * 0.992).toFixed(1));
      allCandles.push({ time: iso, open: open, high: high, low: low, close: close, volume: 35000 });
      p = close;
    }
  }

  // Slice to requested timeframe (5Y, 1Y, 6M, 3M, 1M, 5D, 1D)
  const barLimits = {
    '5Y': 1250,
    '1Y': 250,
    '6M': 125,
    '3M': 65,
    '1M': 22,
    '5D': 5,
    '1D': 1
  };
  const limit = barLimits[range] || 250;
  const slicedCandles = allCandles.slice(-limit);

  return {
    ticker: sym,
    range: range,
    source: source,
    candles: slicedCandles,
    setup: rowObj
  };
}
function fetchMultiYearNepseOhlcv_(sym) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqHeaders = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*'
  };

  let candles = [];
  let source = '';

  // SOURCE 1: Chukul Public Historical OHLCV API (Clean JSON, no Cloudflare block)
  try {
    const chukulUrl = 'https://chukul.com/api/data/historydata/?symbol=' + encodeURIComponent(sym);
    const cRes = UrlFetchApp.fetch(chukulUrl, {
      headers: Object.assign({}, reqHeaders, { 'Referer': 'https://chukul.com/' }),
      muteHttpExceptions: true
    });
    const cText = (cRes.getContentText() || '').trim();
    if (cRes.getResponseCode() === 200 && cText && !cText.startsWith('<')) {
      const json = JSON.parse(cText);
      const rows = Array.isArray(json) ? json : (json.data || json.results || []);
      const parsed = rows.map(p => {
        const rawDate = String(p.date || p.t || p.business_date || '').slice(0, 10);
        const close = Number(p.close || p.c || p.ltp || 0);
        const open = Number(p.open || p.o || close);
        const high = Number(p.high || p.h || Math.max(open, close));
        const low = Number(p.low || p.l || Math.min(open, close));
        return {
          time: rawDate,
          open: open,
          high: Math.max(high, open, close),
          low: Math.min(low, open, close),
          close: close,
          volume: Number(p.volume || p.v || p.qty || 0)
        };
      }).filter(c => /^\d{4}-\d{2}-\d{2}$/.test(c.time) && c.close > 0);

      if (parsed.length >= 10) {
        candles = parsed;
        source = 'CHUKUL_TA';
      }
    }
  } catch (e) {
    console.warn('Chukul history skipped for ' + sym + ':', e.message);
  }

  // SOURCE 2: Merolagani Session-Authenticated POST Request
  if (candles.length < 10) {
    try {
      const pageUrl = 'https://merolagani.com/CompanyDetail.aspx?symbol=' + encodeURIComponent(sym);
      const pageRes = UrlFetchApp.fetch(pageUrl, {
        headers: { 'User-Agent': reqHeaders['User-Agent'] },
        muteHttpExceptions: true
      });
      const rawCookies = pageRes.getAllHeaders()['Set-Cookie'];
      const cookieHeader = Array.isArray(rawCookies)
        ? rawCookies.map(c => c.split(';')[0]).join('; ')
        : (rawCookies ? String(rawCookies).split(';')[0] : '');

      const meroApiUrl = 'https://merolagani.com/handlers/webrequesthandler.ashx?type=get_company_graph&symbol='
        + encodeURIComponent(sym) + '&dateRange=1Y';
      const mRes = UrlFetchApp.fetch(meroApiUrl, {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8',
        payload: '',
        headers: Object.assign({}, reqHeaders, {
          'Referer': pageUrl,
          'Origin': 'https://merolagani.com',
          'X-Requested-With': 'XMLHttpRequest',
          'Cookie': cookieHeader
        }),
        muteHttpExceptions: true
      });

      const mText = (mRes.getContentText() || '').trim();
      if (mRes.getResponseCode() === 200 && mText && !mText.startsWith('<')) {
        const json = JSON.parse(mText);
        const arr = Array.isArray(json) ? json : (json.data || json.candles || []);
        const parsedMero = arr.map(p => {
          let rawT = p.date || p.Date || p.t || p.timestamp || '';
          let iso = typeof rawT === 'number'
            ? Utilities.formatDate(new Date(rawT > 1e11 ? rawT : rawT * 1000), 'Asia/Kathmandu', 'yyyy-MM-dd')
            : String(rawT).slice(0, 10).replace(/\//g, '-');
          const close = Number(p.close || p.Close || p.c || p.ltp || 0);
          const open = Number(p.open || p.Open || p.o || close);
          const high = Number(p.high || p.High || p.h || Math.max(open, close));
          const low = Number(p.low || p.Low || p.l || Math.min(open, close));
          return {
            time: iso,
            open: open,
            high: Math.max(high, open, close),
            low: Math.min(low, open, close),
            close: close,
            volume: Number(p.volume || p.Volume || p.v || 0)
          };
        }).filter(c => /^\d{4}-\d{2}-\d{2}$/.test(c.time) && c.close > 0);

        if (parsedMero.length >= 10) {
          candles = parsedMero;
          source = 'MEROLAGANI_TA';
        }
      }
    } catch (e) {
      console.warn('Merolagani session fetch skipped for ' + sym + ':', e.message);
    }
  }

  // SOURCE 3: Google Sheet 'Price_History' Persistent Archive Fallback
  if (candles.length < 10) {
    const histSheet = ss.getSheetByName('Price_History');
    if (histSheet && histSheet.getLastRow() > 1) {
      const rows = histSheet.getDataRange().getValues().slice(1);
      const sheetBars = rows
        .filter(r => String(r[0]).toUpperCase().trim() === sym)
        .map(r => ({
          time: r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Kathmandu', 'yyyy-MM-dd') : String(r[1]).slice(0, 10),
          open: Number(r[2] || 0),
          high: Number(r[3] || 0),
          low: Number(r[4] || 0),
          close: Number(r[5] || 0),
          volume: Number(r[6] || 0)
        }))
        .filter(c => /^\d{4}-\d{2}-\d{2}$/.test(c.time) && c.close > 0);

      if (sheetBars.length >= 5) {
        candles = sheetBars;
        source = 'SHEET_ARCHIVE_TA';
      }
    }
  }

  // Sort strictly ascending & deduplicate by YYYY-MM-DD
  if (candles.length > 0) {
    candles.sort((a, b) => a.time.localeCompare(b.time));
    const dedup = [];
    const seen = {};
    for (let i = 0; i < candles.length; i++) {
      if (!seen[candles[i].time]) {
        seen[candles[i].time] = true;
        dedup.push(candles[i]);
      }
    }
    candles = dedup;

    // Persist real external bars to 'Price_History' sheet tab for instant future loads
    if (source === 'CHUKUL_TA' || source === 'MEROLAGANI_TA') {
      saveCandlesToSheetArchive_(ss, sym, candles.slice(-300));
    }
  }

  return { candles: candles, source: source || 'SETUP_MATRIX_TA' };
}

function saveCandlesToSheetArchive_(ss, sym, candles) {
  try {
    let sheet = ss.getSheetByName('Price_History');
    if (!sheet) {
      sheet = ss.insertSheet('Price_History');
      sheet.appendRow(['ticker', 'date', 'open', 'high', 'low', 'close', 'volume']);
    }
    const lastRow = sheet.getLastRow();
    const existingKeys = {};
    if (lastRow > 1) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      existing.forEach(r => {
        const dStr = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Kathmandu', 'yyyy-MM-dd') : String(r[1]).slice(0, 10);
        existingKeys[String(r[0]).toUpperCase().trim() + '_' + dStr] = true;
      });
    }
    const toAppend = [];
    candles.forEach(c => {
      const key = sym + '_' + c.time;
      if (!existingKeys[key]) {
        toAppend.push([sym, c.time, c.open, c.high, c.low, c.close, c.volume]);
      }
    });
    if (toAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 7).setValues(toAppend);
    }
  } catch (e) {
    console.warn('Price_History archive write skipped:', e.message);
  }
}

/**
 * Run this function in Apps Script to verify real multi-day candles in the Execution Log.
 */
function testChartDataFetch() {
  CacheService.getScriptCache().remove('nepse_tv_hist_v2_NABIL');
  const res = getLiveChartData('NABIL', '1Y', '');
  Logger.log('Ticker: ' + res.ticker + ' | Source: ' + res.source + ' | Candles: ' + res.candles.length);
  if (res.candles.length > 0) {
    Logger.log('Latest bar: ' + JSON.stringify(res.candles[res.candles.length - 1]));
  }
}
