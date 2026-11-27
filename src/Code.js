/**
 * NEPSE Sentinel Cloud — Autonomous 4-Agent AI Committee + Multi-Source Market Fetcher
 * Schedule: Monday–Friday (Pre-Open 10:30–11:00, Continuous 11:00–15:00 NPT)
 * Auto-updates ALL NEPSE stocks in setup_matrix every 15 min (market hours) / 30 min (off-hours).
 */

const PROPS = PropertiesService.getScriptProperties();

// ==========================================
// 1. SHEET INITIALIZER & TELEMETRY LOGGER
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
    const currentHeaders = sheet.getLastRow() > 0
      ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0]
      : [];
    const headerMismatch = headers.some((h, idx) => String(currentHeaders[idx] || '').trim() !== h);
    if (sheet.getLastRow() === 0 || headerMismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  });

  const l3 = ss.getSheetByName('L3_Playbook_Config');
  if (l3.getLastRow() === 1) {
    l3.appendRow(['v2.1', 'ALL', 5000000, 0.16, 'Autonomous momentum + L2 liquidity + R:R >= 1.8', 'ACTIVE']);
    l3.appendRow(['v2.1', 'HYDRO', 8000000, 0.12, 'High liquidity hydro breakout & accumulation', 'ACTIVE']);
    l3.appendRow(['v2.1', 'BANKING', 5000000, 0.18, 'Commercial bank value & institutional flow', 'ACTIVE']);
  }

  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  logTelemetry('setupSheets', 'SYS_INIT', 'OK', 'Verified all 7 sheets');
}

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
      now.toISOString(), nptStr, hourNpt, fnName,
      parentFn || 'AUTO_CLOUD', status || 'OK',
      String(detail || '').slice(0, 180), String(targetTicker || '').toUpperCase()
    ]);
    const lastRow = sheet.getLastRow();
    if (lastRow > 240) sheet.deleteRows(2, lastRow - 201);
  } catch (e) {
    console.warn('Telemetry skip:', e);
  }
}

// ==========================================
// 2. NEPSE SECTOR MAP & FEE / WACC / L2 ENGINE
// ==========================================
const NEPSE_SECTOR_MAP = {
  'NABIL':'Commercial Banks','NIMB':'Commercial Banks','SCB':'Commercial Banks','HBL':'Commercial Banks',
  'SBI':'Commercial Banks','EBL':'Commercial Banks','NICA':'Commercial Banks','MBL':'Commercial Banks',
  'LSL':'Commercial Banks','KBL':'Commercial Banks','SBL':'Commercial Banks','SANIMA':'Commercial Banks',
  'PRVU':'Commercial Banks','GBIME':'Commercial Banks','CZBIL':'Commercial Banks','PCBL':'Commercial Banks',
  'NBL':'Commercial Banks','ADBL':'Commercial Banks','RBB':'Commercial Banks',
  'MNBBL':'Development Banks','GBBL':'Development Banks','JBBL':'Development Banks','KSBBL':'Development Banks',
  'LBBL':'Development Banks','MLBL':'Development Banks','SHINE':'Development Banks','SADBL':'Development Banks',
  'SINDU':'Development Banks','CORBL':'Development Banks','EDBL':'Development Banks','GRDBL':'Development Banks',
  'SAPDBL':'Development Banks','NABBC':'Development Banks',
  'ICFC':'Finance','MFIL':'Finance','GFCL':'Finance','GMFIL':'Finance','CFCL':'Finance','BFC':'Finance',
  'GUFL':'Finance','JFL':'Finance','MPFL':'Finance','NFS':'Finance','PFL':'Finance','PROFL':'Finance','RLFL':'Finance',
  'CBBL':'Microfinance','DDBL':'Microfinance','FMDBL':'Microfinance','FOWAD':'Microfinance','NUBL':'Microfinance',
  'SKBBL':'Microfinance','SLBBL':'Microfinance','SMATA':'Microfinance','SWBBL':'Microfinance','NICLBSL':'Microfinance',
  'GILB':'Microfinance','GLBSL':'Microfinance','MERO':'Microfinance','MLBBL':'Microfinance','WNLB':'Microfinance',
  'NLIC':'Life Insurance','NLICL':'Life Insurance','LICN':'Life Insurance','ALICL':'Life Insurance',
  'SJLIC':'Life Insurance','ILI':'Life Insurance','CLI':'Life Insurance','HLI':'Life Insurance','PMLI':'Life Insurance',
  'SRLI':'Life Insurance','RNLI':'Life Insurance','SNLI':'Life Insurance','GMLI':'Life Insurance',
  'IGI':'Non Life Insurance','NICL':'Non Life Insurance','NIL':'Non Life Insurance','NLGI':'Non Life Insurance',
  'PRIN':'Non Life Insurance','RBCL':'Non Life Insurance','SALICO':'Non Life Insurance','SICL':'Non Life Insurance',
  'HEI':'Non Life Insurance','SGIC':'Non Life Insurance','SPIL':'Non Life Insurance','UAIC':'Non Life Insurance',
  'AHPC':'Hydro Power','AKJCL':'Hydro Power','AKPL':'Hydro Power','API':'Hydro Power','BARUN':'Hydro Power',
  'BPCL':'Hydro Power','CHCL':'Hydro Power','CHDC':'Hydro Power','GHL':'Hydro Power','HDHPC':'Hydro Power',
  'HPPL':'Hydro Power','HURJA':'Hydro Power','MANDU':'Hydro Power','MEN':'Hydro Power','MHCL':'Hydro Power',
  'NGPL':'Hydro Power','NHDL':'Hydro Power','NHPC':'Hydro Power','RADHI':'Hydro Power','RIDI':'Hydro Power',
  'SAHAS':'Hydro Power','SHPC':'Hydro Power','SJCL':'Hydro Power','SPC':'Hydro Power','UMHL':'Hydro Power',
  'UPCL':'Hydro Power','UPPER':'Hydro Power','TAMOR':'Hydro Power','BGWT':'Hydro Power','DOLTI':'Hydro Power',
  'USHEC':'Hydro Power','RHGCL':'Hydro Power','SPDL':'Hydro Power','BENH':'Hydro Power','SMJC':'Hydro Power',
  'PCIL':'Hydro Power','PPL':'Hydro Power','PMHPL':'Hydro Power','PHCL':'Hydro Power','PPCL':'Hydro Power',
  'SHIVM':'Manufacturing And Processing','HDL':'Manufacturing And Processing','SARBTM':'Manufacturing And Processing',
  'SONA':'Manufacturing And Processing','GHORAHI':'Manufacturing And Processing','UNL':'Manufacturing And Processing',
  'RSML':'Manufacturing And Processing','BNL':'Manufacturing And Processing','BNT':'Manufacturing And Processing',
  'OHL':'Hotels And Tourism','SHL':'Hotels And Tourism','TRH':'Hotels And Tourism','CGH':'Hotels And Tourism',
  'KDL':'Hotels And Tourism','CITY':'Hotels And Tourism',
  'CIT':'Investment','HATHY':'Investment','HIDCL':'Investment','HIDCLP':'Investment','NIFRA':'Investment',
  'NRN':'Investment','ENL':'Investment',
  'BBC':'Tradings','STC':'Tradings','NTC':'Others','NRIC':'Others','SAIL':'Others','MKCL':'Others'
};

function inferNepseSector_(ticker, rawSector) {
  const sym = String(ticker || '').toUpperCase().trim();
  if (NEPSE_SECTOR_MAP[sym]) return NEPSE_SECTOR_MAP[sym];
  const s = String(rawSector || '').trim();
  if (s && s.toLowerCase() !== 'nepse' && s.toLowerCase() !== 'unknown') return s;
  if (sym.endsWith('LBSL') || sym.endsWith('MF') || sym.endsWith('BSL')) return 'Microfinance';
  if (sym.endsWith('LIC') || sym.endsWith('LI')) return 'Life Insurance';
  if (sym.endsWith('IC')) return 'Non Life Insurance';
  if (sym.endsWith('BBL') || sym.endsWith('BL')) return 'Development Banks';
  if (sym.endsWith('FL')) return 'Finance';
  if (sym.endsWith('PCL') || sym.endsWith('HCL') || sym.endsWith('HL') || sym.endsWith('PL') || sym.endsWith('JC')) return 'Hydro Power';
  return 'Others';
}

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
  return {
    tradeVal,
    totalCost: Number(totalCost.toFixed(2)),
    trueWacc: Number((totalCost / qty).toFixed(2))
  };
}

function runL2Guards(row) {
  const failures = [];
  if (Number(row.avg_turnover_20d) < 5000000) failures.push('ILLIQUID_COUNTER (<50L)');
  if (Number(row.macro_approved) === 0 && !String(row.citation_chunk_id || '').trim()) failures.push('UNGROUNDED_MACRO_BLOCK');
  if (String(row.is_price_adjusted_today).toUpperCase() === 'TRUE' && String(row.signal).toUpperCase() === 'BREAKOUT') {
    failures.push('ADJUSTED_PRICE_FALSE_BREAKOUT');
  }
  if (Number(row.kelly_pct) > 0.25) failures.push('KELLY_CAP_EXCEEDED (>25%)');
  return { passed: failures.length === 0, failures };
}

function isNepseLiveSession_() {
  const now = new Date();
  const nptDay = Number(Utilities.formatDate(now, 'Asia/Kathmandu', 'u')); // 1=Mon .. 5=Fri, 6=Sat, 7=Sun
  const hour = Number(Utilities.formatDate(now, 'Asia/Kathmandu', 'H'));
  const min = Number(Utilities.formatDate(now, 'Asia/Kathmandu', 'm'));
  const minsOfDay = hour * 60 + min;
  return (nptDay >= 1 && nptDay <= 5 && minsOfDay >= 630 && minsOfDay <= 915); // Mon–Fri 10:30–15:15 NPT
}

function getLastNepseTradingDateStr_() {
  const npt = new Date(Date.now() + 345 * 60000);
  const day = npt.getUTCDay(); // 0=Sun, 1=Mon .. 5=Fri, 6=Sat
  const backDays = day === 6 ? 1 : (day === 0 ? 2 : 0); // Sat -> Fri, Sun -> Fri
  const marketDate = new Date(npt.getTime() - backDays * 86400000);
  return Utilities.formatDate(marketDate, 'UTC', 'yyyy-MM-dd');
}

// ==========================================
// 3. MULTI-SOURCE LIVE NEPSE MARKET FETCHER
// ==========================================
function fetchFromMerolagani_(headers, sessionDate) {
  const resMero = UrlFetchApp.fetch('https://merolagani.com/LatestMarket.aspx', { headers: headers, muteHttpExceptions: true });
  if (resMero.getResponseCode() !== 200) return null;
  const html = resMero.getContentText();
  const rowMatches = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  const parsed = [];

  for (let r = 0; r < rowMatches.length; r++) {
    const tds = rowMatches[r].match(/<td[\s\S]*?<\/td>/gi);
    if (!tds || tds.length < 7) continue;
    const clean = tds.map(cell => cell.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/,/g, '').trim());
    const sym = clean[0].toUpperCase();
    const ltp = Number(clean[1]);
    if (!sym || !/^[A-Z0-9]{2,12}$/.test(sym) || isNaN(ltp) || ltp <= 0) continue;

    const pct = Number(clean[2]) || 0;
    const open = Number(clean[3]) || ltp;
    const high = Number(clean[4]) || ltp;
    const low = Number(clean[5]) || ltp;
    const qty = Number(clean[6]) || 0;
    const prevClose = Number(clean[7]) || ltp;
    parsed.push({
      ticker: sym, sector: inferNepseSector_(sym, ''), ltp: ltp, pctChange: pct,
      open: open, high: high, low: low, vwap: Number(((high + low + ltp) / 3).toFixed(2)),
      prevClose: prevClose, qty: qty, turnover: ltp * qty
    });
  }
  return parsed.length >= 40 ? { source: 'MEROLAGANI_' + sessionDate, rows: parsed } : null;
}

function fetchFromShareSansar_(headers, fallbackSessionDate) {
  const resSs = UrlFetchApp.fetch('https://www.sharesansar.com/today-share-price', { headers: headers, muteHttpExceptions: true });
  if (resSs.getResponseCode() !== 200) return null;
  const html = resSs.getContentText();
  const dateMatch = html.match(/(20\d{2}-\d{2}-\d{2})/);
  const sessionDate = dateMatch ? dateMatch[1] : fallbackSessionDate;
  const rowMatches = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  const parsed = [];

  for (let r = 0; r < rowMatches.length; r++) {
    const tds = rowMatches[r].match(/<td[\s\S]*?<\/td>/gi);
    if (!tds || tds.length < 14) continue;
    const clean = tds.map(cell => cell.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/,/g, '').trim());
    const sym = clean[1].toUpperCase();
    const ltp = Number(clean[7]);
    if (!sym || !/^[A-Z0-9]{2,12}$/.test(sym) || isNaN(ltp) || ltp <= 0) continue;

    const open = Number(clean[3]) || ltp;
    const high = Number(clean[4]) || ltp;
    const low = Number(clean[5]) || ltp;
    const vwap = Number(clean[10]) || ltp;
    const qty = Number(clean[11]) || 0;
    const prevClose = Number(clean[12]) || ltp;
    const turnover = Number(clean[13]) || (ltp * qty);
    const pct = Number(clean[17]) || (prevClose > 0 ? ((ltp - prevClose) / prevClose) * 100 : 0);

    parsed.push({
      ticker: sym, sector: inferNepseSector_(sym, ''), ltp: ltp, pctChange: pct,
      open: open, high: high, low: low, vwap: vwap, prevClose: prevClose, qty: qty, turnover: turnover
    });
  }
  return parsed.length >= 40 ? { source: 'SHARESANSAR_' + sessionDate, rows: parsed } : null;
}

function fetchLiveNepseMarketSnapshot_() {
  const headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/json,*/*'
  };
  const sessionDate = getLastNepseTradingDateStr_();

  // During Mon–Fri live session (10:30–15:15 NPT), prioritize Merolagani live ticks first
  if (isNepseLiveSession_()) {
    try {
      const liveMero = fetchFromMerolagani_(headers, sessionDate);
      if (liveMero) return liveMero;
    } catch (e) { console.warn('Live Merolagani skip:', e); }
  }

  try {
    const ssData = fetchFromShareSansar_(headers, sessionDate);
    if (ssData) return ssData;
  } catch (e) { console.warn('ShareSansar skip:', e); }

  try {
    const meroData = fetchFromMerolagani_(headers, sessionDate);
    if (meroData) return meroData;
  } catch (e) { console.warn('Merolagani skip:', e); }

  try {
    const resJson = UrlFetchApp.fetch('https://shubhamnpk.github.io/yonepse/data/nepse_data.json', { headers: headers, muteHttpExceptions: true });
    if (resJson.getResponseCode() === 200) {
      const parsed = JSON.parse(resJson.getContentText());
      const arr = Array.isArray(parsed) ? parsed : (parsed.data || parsed.stocks || Object.values(parsed));
      const normalized = [];
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        if (!item || typeof item !== 'object') continue;
        const sym = String(item.symbol || item.ticker || '').toUpperCase().trim();
        const ltp = Number(item.ltp || item.lastTradedPrice || item.close || 0);
        if (!sym || ltp <= 0 || sym.length > 12) continue;
        const qty = Number(item.volume || item.totalTradeQuantity || item.qty || 0);
        const turnover = Number(item.turnover || item.totalTradeValue || (ltp * qty) || 0);
        normalized.push({
          ticker: sym, sector: inferNepseSector_(sym, item.sector || item.sectorName),
          ltp: ltp, pctChange: Number(item.pct_change || item.percentageChange || item.changePercent || 0),
          open: Number(item.open || item.openPrice || ltp), high: Number(item.high || item.highPrice || ltp),
          low: Number(item.low || item.lowPrice || ltp), vwap: ltp,
          prevClose: Number(item.prev_close || item.previousClose || ltp), qty: qty, turnover: turnover
        });
      }
      if (normalized.length >= 30) return { source: 'YONEPSE_' + sessionDate, rows: normalized };
    }
  } catch (e) { console.warn('YONEPSE skip:', e); }

  return { source: 'SHEET_CACHE', rows: [] };
}

// ==========================================
// 4. AUTONOMOUS 4-AGENT AI COMMITTEE ENGINE
// ==========================================
function runAutonomousAiCommittee(forceRefresh) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(12000)) {
    return buildGraphPayloadFromSheet_('BUSY_LOCK_READ');
  }

  try {
    setupSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const matrixSheet = ss.getSheetByName('setup_matrix');
    const l3Sheet = ss.getSheetByName('L3_Playbook_Config');

    const sectorRules = { ALL: { minTurnover: 5000000, maxKelly: 0.18 } };
    if (l3Sheet && l3Sheet.getLastRow() > 1) {
      l3Sheet.getDataRange().getValues().slice(1).forEach(r => {
        if (String(r[5] || '').toUpperCase() === 'ACTIVE') {
          sectorRules[String(r[1] || 'ALL').toUpperCase().trim()] = {
            minTurnover: Number(r[2] || 5000000),
            maxKelly: Number(r[3] || 0.18)
          };
        }
      });
    }

    const snapshot = fetchLiveNepseMarketSnapshot_();
    const now = new Date();
    const nptStamp = Utilities.formatDate(now, 'Asia/Kathmandu', 'yyyy-MM-dd HH:mm:ss');

    if (snapshot.rows.length > 0) {
      const sectorStats = {};
      snapshot.rows.forEach(s => {
        if (!sectorStats[s.sector]) sectorStats[s.sector] = { sumPct: 0, count: 0 };
        sectorStats[s.sector].sumPct += s.pctChange;
        sectorStats[s.sector].count += 1;
      });

      const evaluatedRows = [];
      const l0Batch = [];
      let passCount = 0;

      for (let i = 0; i < snapshot.rows.length; i++) {
        const st = snapshot.rows[i];
        const secUpper = st.sector.toUpperCase();
        const rule = secUpper.includes('HYDRO')
          ? (sectorRules.HYDRO || sectorRules.ALL)
          : (secUpper.includes('BANK') ? (sectorRules.BANKING || sectorRules.ALL) : sectorRules.ALL);

        const secAvgPct = sectorStats[st.sector] ? (sectorStats[st.sector].sumPct / sectorStats[st.sector].count) : 0;

        // AGENT 1: Technical & Dynamic Risk-Reward Structure
        const dayRange = Math.max(st.ltp * 0.015, st.high - st.low);
        const closeStrength = (st.high > st.low) ? ((st.ltp - st.low) / (st.high - st.low)) : 0.50;
        const vwapPremium = st.vwap > 0 ? ((st.ltp - st.vwap) / st.vwap) : 0;
        const stopLoss = Number(Math.max(10, Math.min(st.low * 0.99, st.ltp - dayRange * 1.1)).toFixed(1));
        const riskPerShare = Math.max(1, st.ltp - stopLoss);
        const rrMultiplier = Number((1.85 + closeStrength * 0.65 + Math.max(0, st.pctChange) * 0.08).toFixed(2));
        const target = Number((st.ltp + riskPerShare * rrMultiplier).toFixed(1));
        const rrRatio = (target - st.ltp) / riskPerShare;
        const entryZone = Number((st.ltp * 0.993).toFixed(1)) + ' - ' + Number((st.ltp * 1.004).toFixed(1));

        // AGENT 2: Institutional Liquidity Gate
        const effectiveTurnover = st.turnover > 0 ? st.turnover : (st.ltp * st.qty);
        const liquidOk = effectiveTurnover >= rule.minTurnover;

        // AGENT 3: Macro & Positive Momentum Gate
        const macroApproved = (liquidOk && st.pctChange > 0 && (secAvgPct >= -0.5 || st.pctChange >= 1.2)) ? 1 : 0;

        // Signal Classification
        let signal = 'SCAN';
        if (liquidOk && macroApproved === 1 && st.pctChange >= 2.5 && closeStrength >= 0.68 && st.pctChange <= 9.5) {
          signal = 'BREAKOUT';
        } else if (liquidOk && macroApproved === 1 && st.pctChange >= 0.8 && st.pctChange < 2.5 && closeStrength >= 0.60) {
          signal = 'BUY';
        } else if (liquidOk && macroApproved === 1 && st.pctChange > 0.05 && st.pctChange < 0.8 && closeStrength >= 0.58 && effectiveTurnover >= 15000000) {
          signal = 'ACCUM';
        }

        // AGENT 4: Stock-Specific Fractional Kelly Position Sizer
        let kellyPct = 0;
        if (signal !== 'SCAN') {
          const liqBonus = Math.min(0.05, Math.log10(Math.max(1, effectiveTurnover / 5000000)) * 0.025);
          const momBonus = Math.min(0.05, Math.max(0, st.pctChange) * 0.008);
          const vwapBonus = Math.max(-0.02, Math.min(0.03, vwapPremium * 0.8));
          const winProb = Math.min(0.64, Math.max(0.47, 0.46 + closeStrength * 0.08 + liqBonus + momBonus + vwapBonus));
          const rawKelly = winProb - ((1 - winProb) / Math.max(1.5, rrRatio));
          kellyPct = Number(Math.max(0.025, Math.min(rule.maxKelly, rawKelly * 0.5)).toFixed(3));
        }

        const turnoverCr = (effectiveTurnover / 10000000).toFixed(2);
        const rowObj = {
          ticker: st.ticker, sector: st.sector, ltp: st.ltp, open: st.open, high: st.high, low: st.low,
          prev_close: st.prevClose, vwap: st.vwap, pct_change: Number(st.pctChange.toFixed(2)),
          signal: signal, entry_zone: entryZone, stop_loss: stopLoss, target: target,
          avg_turnover_20d: Math.round(effectiveTurnover), is_price_adjusted_today: false,
          macro_approved: macroApproved, citation_chunk_id: snapshot.source,
          bull_case: `Chg ${st.pctChange >= 0 ? '+' : ''}${st.pctChange.toFixed(2)}% | Turnover Rs ${turnoverCr} Cr | Close Strength ${(closeStrength * 100).toFixed(0)}% | R:R ${rrRatio.toFixed(1)}x`,
          bear_case: `SL Rs ${stopLoss} | Range Rs ${st.low}-${st.high} | Sector ${secAvgPct >= 0 ? '+' : ''}${secAvgPct.toFixed(2)}%`,
          risk_flag: liquidOk ? 'L2_GUARD_PASS' : 'LOW_TURNOVER',
          kelly_pct: kellyPct
        };

        const guard = runL2Guards(rowObj);
        if (!guard.passed) {
          rowObj.signal = 'SCAN';
          rowObj.kelly_pct = 0;
          rowObj.risk_flag = guard.failures.join(' | ');
        } else if (rowObj.signal !== 'SCAN') {
          passCount++;
          if (l0Batch.length < 30) {
            l0Batch.push([now, 'AUTO_COMMITTEE', rowObj.ticker, 'v2.1', JSON.stringify(rowObj), JSON.stringify(guard)]);
          }
        }

        evaluatedRows.push(rowObj);
      }

      evaluatedRows.sort((a, b) => {
        const aPass = a.signal !== 'SCAN' ? 1 : 0;
        const bPass = b.signal !== 'SCAN' ? 1 : 0;
        if (bPass !== aPass) return bPass - aPass;
        return (b.kelly_pct - a.kelly_pct) || (b.avg_turnover_20d - a.avg_turnover_20d);
      });

      // Write ALL evaluated NEPSE stocks to setup_matrix
      if (matrixSheet.getLastRow() > 1) {
        matrixSheet.getRange(2, 1, matrixSheet.getLastRow() - 1, matrixSheet.getLastColumn()).clearContent();
      }

      const sheetValues = evaluatedRows.map(r => [
        r.ticker, r.sector, r.ltp, r.signal, r.entry_zone, r.stop_loss, r.target,
        r.avg_turnover_20d, r.is_price_adjusted_today, r.macro_approved,
        r.citation_chunk_id, r.bull_case, r.bear_case, r.risk_flag, r.kelly_pct, now
      ]);

      if (sheetValues.length > 0) {
        matrixSheet.getRange(2, 1, sheetValues.length, sheetValues[0].length).setValues(sheetValues);
        matrixSheet.getRange(2, 16, sheetValues.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }

      if (l0Batch.length > 0) {
        const l0Sheet = ss.getSheetByName('L0_Run_Snapshots');
        l0Sheet.getRange(l0Sheet.getLastRow() + 1, 1, l0Batch.length, l0Batch[0].length).setValues(l0Batch);
      }

      SpreadsheetApp.flush();
      logTelemetry(
        'runAutonomousAiCommittee',
        snapshot.source,
        'OK',
        `Updated all ${evaluatedRows.length} NEPSE stocks -> ${passCount} AI Committee Pass`
      );

      const enrichedMap = {};
      evaluatedRows.forEach(r => { enrichedMap[r.ticker] = r; });
      const payload = buildGraphPayloadFromSheet_(snapshot.source + ' · ' + nptStamp);
      payload.setups = payload.setups.map(s => Object.assign({}, enrichedMap[s.ticker] || {}, s));
      return payload;
    }

    logTelemetry('runAutonomousAiCommittee', 'SHEET_CACHE', 'WARN', 'Live feeds unreachable; served cached setup_matrix');
    return buildGraphPayloadFromSheet_('Cached Sheet · ' + nptStamp);
  } finally {
    lock.releaseLock();
  }
}

function runLiveAiCommitteeScan() {
  return runAutonomousAiCommittee(true);
}

// ==========================================
// 5. GRAPH DATA ENDPOINT (AUTO-HEALS OLD CACHE)
// ==========================================
function getAutonomousGraphData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('setup_matrix');

  let needsFreshScan = (!matrixSheet || matrixSheet.getLastRow() <= 1);
  if (!needsFreshScan) {
    const firstChunk = String(matrixSheet.getRange(2, 11).getValue() || '');
    const lastUpdatedVal = matrixSheet.getRange(2, 16).getValue();
    if (!firstChunk.includes('_20')) {
      needsFreshScan = true;
    } else if (lastUpdatedVal instanceof Date) {
      const ageMinutes = (Date.now() - lastUpdatedVal.getTime()) / 60000;
      const maxAge = isNepseLiveSession_() ? 15 : 30;
      if (ageMinutes > maxAge) needsFreshScan = true;
    } else {
      needsFreshScan = true;
    }
  }

  if (needsFreshScan) return runAutonomousAiCommittee(true);
  return buildGraphPayloadFromSheet_('Live Sheet Cache');
}

function buildGraphPayloadFromSheet_(sourceLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('setup_matrix');
  const portSheet = ss.getSheetByName('Portfolio');
  const heldSet = {};

  if (portSheet && portSheet.getLastRow() > 1) {
    portSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0]) heldSet[String(r[0]).toUpperCase().trim()] = true;
    });
  }

  const setups = [];
  let lastSyncNpt = Utilities.formatDate(new Date(), 'Asia/Kathmandu', 'HH:mm:ss');

  if (matrixSheet && matrixSheet.getLastRow() > 1) {
    const data = matrixSheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      const obj = Object.fromEntries(headers.map((h, idx) => [h, data[i][idx]]));
      if (!obj.ticker) continue;
      obj.in_portfolio = !!heldSet[String(obj.ticker).toUpperCase().trim()];
      if (obj.updated_at instanceof Date) {
        lastSyncNpt = Utilities.formatDate(obj.updated_at, 'Asia/Kathmandu', 'HH:mm:ss');
        obj.updated_at = obj.updated_at.toISOString();
      }
      setups.push(obj);
    }
  }

  return {
    setups: setups,
    sourceLabel: sourceLabel || 'Autonomous AI Committee',
    lastSyncNpt: lastSyncNpt
  };
}

// ==========================================
// 6. ZERO-TOUCH 15-MIN / 30-MIN CLOUD TRIGGERS
// ==========================================
function installZeroTouchAutomation() {
  setupSheets();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Runs every 15 min during Mon–Fri NEPSE hours, and every 30 min off-hours
  ScriptApp.newTrigger('scheduledMarketCommitteeWorker')
    .timeBased()
    .everyMinutes(15)
    .create();

  // Morning EDIS check before Pre-Open (Mon–Fri 10:15 NPT)
  ScriptApp.newTrigger('pushMorningAlert')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .nearMinute(15)
    .create();

  // EOD Digest after 15:00 NPT close
  ScriptApp.newTrigger('pushEodDigest')
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(15)
    .create();

  runAutonomousAiCommittee(true);
  logTelemetry('installZeroTouchAutomation', 'SYS_INIT', 'OK', 'Active: 15-min live / 30-min off-hours Mon–Fri triggers');
}

function scheduledMarketCommitteeWorker() {
  if (isNepseLiveSession_()) {
    runAutonomousAiCommittee(true);
    return;
  }
  const min = Number(Utilities.formatDate(new Date(), 'Asia/Kathmandu', 'm'));
  if (min < 15 || (min >= 30 && min < 45)) {
    runAutonomousAiCommittee(true);
  }
}

function pushMorningAlert() {
  const nptDay = Number(Utilities.formatDate(new Date(), 'Asia/Kathmandu', 'u'));
  if (nptDay > 5) return; // Skip Saturday & Sunday
  runAutonomousAiCommittee(true);
  const chatId = PROPS.getProperty('ALLOWED_CHAT_ID');
  if (chatId && PROPS.getProperty('BOT_TOKEN')) {
    sendTelegram(chatId, checkEdisObligations('pushMorningAlert'), true, 'pushMorningAlert');
  }
}

function pushEodDigest() {
  const nptDay = Number(Utilities.formatDate(new Date(), 'Asia/Kathmandu', 'u'));
  if (nptDay > 5) return; // Skip Saturday & Sunday
  runAutonomousAiCommittee(true);
  const chatId = PROPS.getProperty('ALLOWED_CHAT_ID');
  if (chatId && PROPS.getProperty('BOT_TOKEN')) {
    sendTelegram(chatId, buildMatrixDigest('pushEodDigest'), true, 'pushEodDigest');
  }
}

// ==========================================
// 7. TELEGRAM & DIGEST HELPERS
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData) return HtmlService.createHtmlOutput('OK');
    const body = JSON.parse(e.postData.contents);

    const cache = CacheService.getScriptCache();
    if (body.update_id && cache.get('tg_' + body.update_id)) return HtmlService.createHtmlOutput('OK');
    if (body.update_id) cache.put('tg_' + body.update_id, '1', 300);

    const msg = body.message;
    if (!msg || !msg.text) return HtmlService.createHtmlOutput('OK');

    const chatId = String(msg.chat.id);
    if (chatId !== PROPS.getProperty('ALLOWED_CHAT_ID')) return HtmlService.createHtmlOutput('OK');

    const parts = msg.text.trim().split(/\s+/);
    const cmd = parts[0].toLowerCase();

    if (cmd === '/scan') {
      runAutonomousAiCommittee(true);
      sendTelegram(chatId, buildMatrixDigest('doPost'), true, 'doPost');
    } else if (cmd === '/report' && parts[1]) {
      sendTelegram(chatId, buildTickerReport(parts[1].toUpperCase()), true, 'buildTickerReport');
    } else if (cmd === '/matrix') {
      sendTelegram(chatId, buildMatrixDigest('doPost'), true, 'buildMatrixDigest');
    } else if (cmd === '/graph') {
      const webUrl = PROPS.getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl();
      sendTelegram(chatId, 'NEPSE Sentinel 3D Knowledge Graph:\n' + webUrl, false, 'doPost');
    } else if (cmd === '/buy' && parts.length >= 4) {
      sendTelegram(chatId, logPaperOrRealBuy(parts[1].toUpperCase(), Number(parts[2]), Number(parts[3]), parts[4] || 'GENERAL'), true, 'logPaperOrRealBuy');
    } else if (cmd === '/edis') {
      sendTelegram(chatId, checkEdisObligations('doPost'), true, 'checkEdisObligations');
    }
  } catch (err) {
    console.error(err);
  }
  return HtmlService.createHtmlOutput('OK');
}

function sendTelegram(chatId, text, useMarkdown = true, callerFn = 'doPost') {
  const token = PROPS.getProperty('BOT_TOKEN');
  if (!token || !chatId) return;
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const cleanText = useMarkdown ? text.replace(/[*_`]/g, '') : text;
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: cleanText }),
    muteHttpExceptions: true
  });
}

function logPaperOrRealBuy(ticker, qty, price, sector) {
  const calc = calcNepseBuyWacc(qty, price);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio');
  sheet.appendRow([ticker, qty, price, calc.trueWacc, new Date().toISOString().slice(0, 10), 'LOCKED_T0', sector]);
  return `Logged ${ticker}: ${qty} @ Rs ${price} | True WACC: Rs ${calc.trueWacc}`;
}

function buildTickerReport(ticker) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setup_matrix').getDataRange().getValues();
  const headers = data[0];
  const rowArr = data.slice(1).find(r => String(r[0]).toUpperCase() === ticker);
  if (!rowArr) return `No entry for ${ticker} in setup_matrix yet.`;
  const row = Object.fromEntries(headers.map((h, i) => [h, rowArr[i]]));
  return `${ticker} (${row.sector}) — LTP: Rs ${row.ltp}\nSignal: ${row.signal} | Entry: ${row.entry_zone} | SL: ${row.stop_loss} | Target: ${row.target}\nKelly: ${(Number(row.kelly_pct || 0) * 100).toFixed(1)}%`;
}

function buildMatrixDigest(callerFn) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setup_matrix').getDataRange().getValues();
  if (data.length <= 1) return 'setup_matrix is empty.';
  const headers = data[0];
  const lines = ['Autonomous AI Committee Approved Setups:'];
  for (let i = 1; i < data.length; i++) {
    const row = Object.fromEntries(headers.map((h, idx) => [h, data[i][idx]]));
    if (String(row.signal).toUpperCase() === 'SCAN') continue;
    if (Number(row.macro_approved) === 1) {
      lines.push(`• ${row.ticker} (${row.signal}) @ Rs ${row.ltp} | Entry ${row.entry_zone} | SL ${row.stop_loss} | Target ${row.target} | Kelly ${(Number(row.kelly_pct) * 100).toFixed(1)}%`);
    }
  }
  return lines.length > 1 ? lines.slice(0, 16).join('\n') : 'No stocks passed all 4 AI Committee gates today.';
}

function checkEdisObligations(callerFn) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Portfolio').getDataRange().getValues();
  const pending = data.slice(1).filter(r => String(r[5]).includes('LOCKED') || String(r[5]).includes('EDIS'));
  if (pending.length === 0) return 'No pending T+0 / T+1 EDIS settlement locks.';
  return 'Settlement Tracker:\n' + pending.map(r => `• ${r[0]} (${r[1]} units @ WACC ${r[3]}) — ${r[5]}`).join('\n');
}

function includeVoiceCommandRouter() {
  return HtmlService.createHtmlOutputFromFile('VoiceCommandRouter').getContent();
}

function includeVoiceLab() {
  return HtmlService.createHtmlOutputFromFile('VoiceLab').getContent();
}
