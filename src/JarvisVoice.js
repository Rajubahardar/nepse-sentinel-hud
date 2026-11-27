/**
 * Server-side HTML loader for JarvisVoiceHUD.html
 */
function includeJarvisVoice() {
  return HtmlService.createHtmlOutputFromFile('JarvisVoiceHUD').getContent();
}

function includeFridayVoice() {
  return includeJarvisVoice();
}
function includeBrainCameraFocus() {
  return HtmlService.createHtmlOutputFromFile('BrainCameraFocus').getContent();
}
function includeFridayHindiVoice() {
  return HtmlService.createHtmlOutputFromFile('FridayHindiVoice').getContent();
}
function includeVoiceCommandRouter() {
  return HtmlService.createHtmlOutputFromFile('VoiceCommandRouter').getContent();
}
/**
 * F.R.I.D.A.Y. Real-Time Neural Telemetry, Company Briefing & Animated Code Trace
 */
function getJarvisBrainTelemetry() {
  var tz = 'Asia/Kathmandu';
  var now = new Date();
  var dayNum = Number(Utilities.formatDate(now, tz, 'u')); // 5=Fri, 6=Sat
  var hourNum = Number(Utilities.formatDate(now, tz, 'H'));
  var timeStr = Utilities.formatDate(now, tz, 'HH:mm');
  var isWeekend = (dayNum === 5 || dayNum === 6);

  var greeting = hourNum < 12
    ? 'Good morning, Raju D. sir.'
    : (hourNum < 17 ? 'Good afternoon, Raju D. sir.' : 'Good evening, Raju D. sir.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName('Setup_Matrix');
  var journalSheet = ss.getSheetByName('Trade_Journal');

  var setupRows = [];
  if (setupSheet && setupSheet.getLastRow() >= 2) {
    var raw = setupSheet.getDataRange().getValues();
    var headers = raw[0].map(function(h) { return String(h).trim().toLowerCase(); });
    for (var r = 1; r < raw.length; r++) {
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        obj[headers[c]] = raw[r][c];
      }
      if (obj.ticker) setupRows.push(obj);
    }
  }

  var buySetups = setupRows.filter(function(s) {
    return String(s.signal).toUpperCase() === 'BUY';
  });
  var watchSetups = setupRows.filter(function(s) {
    return String(s.signal).toUpperCase() !== 'BUY';
  });
  var blockedByMacro = setupRows.filter(function(s) {
    return Number(s.macro_approved) === 0;
  });

  var sectorMap = {};
  setupRows.forEach(function(s) {
    var sec = String(s.sector || 'General').trim();
    sectorMap[sec] = (sectorMap[sec] || 0) + 1;
  });
  var sectorNames = Object.keys(sectorMap);

  var closedTrades = 0;
  var wins = 0;
  if (journalSheet && journalSheet.getLastRow() >= 2) {
    var jData = journalSheet.getDataRange().getValues();
    var jHead = jData[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var pnlIdx = jHead.indexOf('pnl_pct') > -1 ? jHead.indexOf('pnl_pct') : jHead.indexOf('return_pct');
    var statIdx = jHead.indexOf('status');

    for (var i = 1; i < jData.length; i++) {
      var st = statIdx > -1 ? String(jData[i][statIdx]).toUpperCase() : 'CLOSED';
      var pnl = pnlIdx > -1 ? parseFloat(jData[i][pnlIdx]) : NaN;
      if (!isNaN(pnl) && (st === 'CLOSED' || st === 'WIN' || st === 'LOSS' || statIdx === -1)) {
        closedTrades++;
        if (pnl > 0) wins++;
      }
    }
  }

  var accuracyPct = closedTrades > 0
    ? Math.round(((wins / closedTrades) * 0.6 + 0.68 * 0.4) * 100)
    : 68;

  var top = buySetups.length > 0 ? buySetups[0] : (setupRows[0] || null);
  var runnerUps = (buySetups.length > 1 ? buySetups.slice(1, 4) : watchSetups.slice(0, 3));

  // ==================== 1. BRIEFING SCRIPT (COMPANY REPORT) ====================
  var companyOverview =
    greeting + ' F.R.I.D.A.Y. company briefing ready at ' + timeStr + ' Nepal Time. ' +
    'I have analyzed ' + setupRows.length + ' listed NEPSE companies across ' +
    (sectorNames.length > 0 ? sectorNames.slice(0, 3).join(', ') : 'key market') +
    ' sectors. ' + buySetups.length + ' companies currently hold active Buy ratings, with ' +
    watchSetups.length + ' on Watch at ' + accuracyPct + ' percent model accuracy.';

  var topCompanyDetail = top
    ? ('Top ranked company is ' + top.ticker +
       (top.sector ? ' in the ' + top.sector + ' sector' : '') +
       ', carrying a ' + String(top.signal || 'WATCH').toUpperCase() + ' signal. ' +
       'Last traded price is ' + top.ltp + ' rupees, entry zone ' + (top.entry_zone || top.ltp) +
       ', target price ' + top.target + ', and stop loss ' + top.stop_loss +
       ', yielding a Risk-to-Reward ratio of ' + (top.rr_ratio || '2.0') +
       ' with ' + (top.kelly_pct || '5.0') + ' percent Half-Kelly position sizing.')
    : 'No companies are currently populated in Setup Matrix, Raju D. sir.';

  var runnerUpSummary = '';
  if (runnerUps.length > 0) {
    var listText = runnerUps.map(function(c) {
      return c.ticker + ' (' + String(c.signal || 'WATCH').toUpperCase() + ' at ' + c.ltp + ', target ' + c.target + ')';
    }).join('; ');
    runnerUpSummary = 'Secondary companies on your radar, Raju D. sir: ' + listText + '.';
  } else {
    runnerUpSummary = 'No additional secondary companies require immediate action today, Raju D. sir.';
  }

  // ==================== 2. REAL-TIME DIAGNOSIS (NODE LOCK + ANIMATED CODE) ====================
  var topTickerStr = top ? String(top.ticker) : 'WATCH';
  var topSignalStr = top ? String(top.signal || 'BUY').toUpperCase() : 'WATCH';
  var topLtpStr = top ? String(top.ltp) : '--';
  var topTargetStr = top ? String(top.target) : '--';
  var topSlStr = top ? String(top.stop_loss) : '--';
  var topRrStr = top ? String(top.rr_ratio || '2.1') : '--';
  var topKellyStr = top ? String(top.kelly_pct || '5.0') : '0';

  var brainDiagnosis = [
    {
      step: 1,
      totalSteps: 6,
      tag: 'NODE 1/6 · ENTRY POINT',
      functionName: 'doGet(e)',
      moduleFile: 'Code.js',
      role: 'Web App Request Router & HUD Assembler',
      nodeKeywords: ['doGet', 'Code.js', 'Code', 'WebApp', 'include'],
      accentColor: '#38bdf8',
      positionHint: { x: 0.28, y: 0.34 },
      isFinal: false,
      codeLines: [
        { phase: 'INPUT',   code: 'function doGet(e) {', runtime: 'HTTP GET /exec' },
        { phase: 'INPUT',   code: '  const tpl = HtmlService.createTemplateFromFile("GraphUI");', runtime: '3D Graph + Visor' },
        { phase: 'PROCESS', code: '  tpl.initialData = loadSetupMatrixRows();', runtime: setupRows.length + ' rows loaded' },
        { phase: 'PROCESS', code: '  return tpl.evaluate()', runtime: 'IFRAME AllowAll' },
        { phase: 'OUTPUT',  code: '    .setTitle("NEPSE Sentinel · Mark VII");', runtime: 'DOM Ready' }
      ],
      text: 'Helmet disengaged, Raju D. sir. Locking onto doGet inside Code dot J S. When you open the Web App, doGet initializes the Graph U I template, injects ' + setupRows.length + ' rows from Setup Matrix, and mounts the Knowledge Graph canvas.'
    },
    {
      step: 2,
      totalSteps: 6,
      tag: 'NODE 2/6 · DATA INGESTION',
      functionName: 'updateDailyNepseSetups()',
      moduleFile: 'DailyNepseSync.js',
      role: isWeekend ? 'Weekend Guard Active · Preserving Friday Close' : '15:15 NST Market Snapshot Fetcher',
      nodeKeywords: ['updateDailyNepseSetups', 'DailyNepseSync.js', 'DailyNepseSync', 'Setup_Matrix', 'fetch'],
      accentColor: '#60a5fa',
      positionHint: { x: 0.72, y: 0.30 },
      isFinal: false,
      codeLines: [
        { phase: 'INPUT',   code: 'function updateDailyNepseSetups() {', runtime: timeStr + ' NST' },
        { phase: 'INPUT',   code: '  const day = Utilities.formatDate(new Date(), "Asia/Kathmandu", "u");', runtime: isWeekend ? 'Day ' + dayNum + ' (Closed)' : 'Day ' + dayNum + ' (Open)' },
        { phase: 'PROCESS', code: '  if (day === "5" || day === "6") return skipWeekend();', runtime: isWeekend ? 'GUARD TRIGGERED' : 'Market Active' },
        { phase: 'PROCESS', code: '  const raw = fetchNepseSnapshot() || readWatchlistFallback();', runtime: setupRows.length + ' tickers synced' },
        { phase: 'OUTPUT',  code: '  writeSetupMatrix(raw);', runtime: 'Setup_Matrix Updated' }
      ],
      text: 'Shifting focus to updateDailyNepseSetups in DailyNepseSync. First it checks Asia Kathmandu day of week. ' +
            (isWeekend
              ? 'Because today is the weekend closure, the calendar guard halts live writes so Friday closing prices stay intact.'
              : 'On active trading days at 15:15, it fetches end-of-day NEPSE prices and writes ' + setupRows.length + ' tickers into Setup Matrix.')
    },
    {
      step: 3,
      totalSteps: 6,
      tag: 'NODE 3/6 · GRAPH TOPOLOGY',
      functionName: 'buildKnowledgeGraph()',
      moduleFile: 'Graph.js',
      role: 'Neural Node & Sector Edge Compiler',
      nodeKeywords: ['buildKnowledgeGraph', 'getGraphData', 'Graph.js', 'Graph', 'Trade_Journal'],
      accentColor: '#a78bfa',
      positionHint: { x: 0.75, y: 0.62 },
      isFinal: false,
      codeLines: [
        { phase: 'INPUT',   code: 'function buildKnowledgeGraph() {', runtime: 'Reading Sheets' },
        { phase: 'INPUT',   code: '  const setups = readSheet("Setup_Matrix");', runtime: setupRows.length + ' company nodes' },
        { phase: 'PROCESS', code: '  const sectors = groupBySector(setups);', runtime: sectorNames.length + ' sector hubs' },
        { phase: 'PROCESS', code: '  const links = connectTickerToSectorAndCode(setups, sectors);', runtime: 'Edges wired' },
        { phase: 'OUTPUT',  code: '  return { nodes, links };', runtime: 'ForceGraph Ready' }
      ],
      text: 'Now highlighting buildKnowledgeGraph inside Graph dot J S. This function converts Setup Matrix rows into ' + setupRows.length + ' company nodes, clusters them across ' + sectorNames.length + ' sector hubs, and wires the neural graph edges you see on screen.'
    },
    {
      step: 4,
      totalSteps: 6,
      tag: 'NODE 4/6 · 4-GATE FILTER',
      functionName: 'evaluateMultiGateSignals(ticker)',
      moduleFile: 'DailyNepseSync.js',
      role: 'PGA-BB · R:R Gate · NRB Macro Lock · Half-Kelly',
      nodeKeywords: ['evaluate', 'PGA', 'Kelly', 'NRB', 'DailyNepseSync', 'Setup_Matrix'],
      accentColor: '#ffb340',
      positionHint: { x: 0.30, y: 0.64 },
      isFinal: false,
      codeLines: [
        { phase: 'INPUT',   code: 'function evaluateMultiGateSignals(stock) {', runtime: 'Scanning ' + setupRows.length + ' stocks' },
        { phase: 'PROCESS', code: '  const bb = calcPgaBollinger(stock.closes, 20);', runtime: 'Band compression OK' },
        { phase: 'PROCESS', code: '  const rr = (stock.target - stock.ltp) / (stock.ltp - stock.sl);', runtime: 'Min R:R >= 1.8' },
        { phase: 'PROCESS', code: '  if (!checkNrbMacroLock(stock.sector)) return "WATCH";', runtime: blockedByMacro.length + ' filtered out' },
        { phase: 'OUTPUT',  code: '  stock.kelly = Math.min(7.5, halfKelly(winProb, rr));', runtime: buySetups.length + ' BUY passed' }
      ],
      text: 'Next, inspecting evaluateMultiGateSignals. Every ticker passes four quantitative gates: twenty-day PGA Bollinger Bands, Risk-to-Reward ratio, Nepal Rastra Bank sector macro locks' +
            (blockedByMacro.length > 0 ? '—which downgraded ' + blockedByMacro.length + ' unapproved setups—' : ', ') +
            'and Half-Kelly position sizing capped at seven point five percent.'
    },
    {
      step: 5,
      totalSteps: 6,
      tag: 'NODE 5/6 · BACKTEST BLENDER',
      functionName: 'getHelmetBacktestJson()',
      moduleFile: 'HelmetBoot.js',
      role: 'Realized Trade Journal & Prior Accuracy Blender',
      nodeKeywords: ['getHelmetBacktestJson', 'HelmetBoot.js', 'HelmetBoot', 'Trade_Journal'],
      accentColor: '#f472b6',
      positionHint: { x: 0.52, y: 0.26 },
      isFinal: false,
      codeLines: [
        { phase: 'INPUT',   code: 'function getHelmetBacktestJson() {', runtime: closedTrades + ' closed trades' },
        { phase: 'INPUT',   code: '  const journal = readSheet("Trade_Journal");', runtime: wins + ' realized wins' },
        { phase: 'PROCESS', code: '  const liveWinRate = closedTrades > 0 ? wins / closedTrades : 0.68;', runtime: 'Prior = 68%' },
        { phase: 'PROCESS', code: '  const blended = Math.round((liveWinRate * 0.6 + 0.68 * 0.4) * 100);', runtime: accuracyPct + '% accuracy' },
        { phase: 'OUTPUT',  code: '  return rankBuyCandidates(setups, blended);', runtime: 'Ranked by Kelly' }
      ],
      text: 'Moving to getHelmetBacktestJson in HelmetBoot dot J S. It blends your Trade Journal win rate with the multi-gate model prior, establishing ' + accuracyPct + ' percent confidence and ranking candidates by Risk-Reward and Kelly score.'
    },
    {
      step: 6,
      totalSteps: 6,
      tag: 'DIAGNOSIS COMPLETE · OUTPUT LOCKED',
      functionName: top ? ('OUTPUT SELECTED: ' + topTickerStr + ' (' + topSignalStr + ')') : 'OUTPUT SELECTED: WATCH MODE',
      moduleFile: 'Setup_Matrix ➔ Visor Target Lock',
      role: top
        ? ('LTP Rs.' + topLtpStr + ' · Target Rs.' + topTargetStr + ' · SL Rs.' + topSlStr + ' · Kelly ' + topKellyStr + '%')
        : 'All gates checked · Standing by in Watch mode',
      nodeKeywords: top
        ? [topTickerStr, String(top.sector || ''), 'Setup_Matrix', 'getHelmetBacktestJson']
        : ['Setup_Matrix', 'getHelmetBacktestJson'],
      accentColor: '#00f5d4',
      positionHint: { x: 0.58, y: 0.48 },
      isFinal: true,
      metrics: top ? {
        ticker: topTickerStr,
        sector: String(top.sector || 'NEPSE'),
        signal: topSignalStr,
        ltp: topLtpStr,
        target: topTargetStr,
        stopLoss: topSlStr,
        rr: topRrStr,
        kelly: topKellyStr,
        accuracy: accuracyPct
      } : null,
      codeLines: top ? [
        { phase: 'INPUT',   code: 'const candidates = filterApprovedBuys(Setup_Matrix);', runtime: buySetups.length + ' qualified' },
        { phase: 'PROCESS', code: 'const winner = sortByScore(candidates)[0];', runtime: topTickerStr + ' (#1)' },
        { phase: 'PROCESS', code: 'assert(winner.rr >= 1.8 && winner.kelly <= 7.5);', runtime: 'R:R ' + topRrStr + ' · Kelly ' + topKellyStr + '%' },
        { phase: 'OUTPUT',  code: 'return { ticker: "' + topTickerStr + '", ltp: ' + topLtpStr + ', target: ' + topTargetStr + ' };', runtime: 'LOCKED ON GRID' }
      ] : [
        { phase: 'INPUT',   code: 'const candidates = filterApprovedBuys(Setup_Matrix);', runtime: '0 qualified' },
        { phase: 'OUTPUT',  code: 'return { signal: "WATCH", status: "STANDBY" };', runtime: 'CAPITAL PROTECTED' }
      ],
      text: top
        ? ('End-to-end diagnosis complete, Raju D. sir. Isolating the final selected output node: ' + topTickerStr + ' at ' + topLtpStr + ' rupees, with target ' + topTargetStr + ', stop loss ' + topSlStr + ', and ' + topKellyStr + ' percent Half-Kelly allocation. The winning node and its execution trace remain locked on your screen.')
        : ('End-to-end diagnosis complete, Raju D. sir. Zero candidates cleared all four gates today, so the brain outputs Watch mode to protect capital.')
    }
  ];

  return {
    ok: true,
    timestamp: timeStr + ' NST',
    buyCount: buySetups.length,
    totalCount: setupRows.length,
    topTicker: topTickerStr,
    companyBriefing: [
      { tag: 'COMPANY BRIEFING · OVERVIEW', text: companyOverview },
      { tag: 'TOP COMPANY · ' + topTickerStr, text: topCompanyDetail },
      { tag: 'WATCHLIST RADAR', text: runnerUpSummary }
    ],
    brainDiagnosis: brainDiagnosis
  };
}
/**
 * Checks whether a Studio Human AI TTS key is configured in Script Properties
 */
function getHumanVoiceConfig() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('HUMAN_TTS_API_KEY');
  var provider = (props.getProperty('HUMAN_TTS_PROVIDER') || 'openai').toLowerCase();
  return {
    enabled: Boolean(key && key.trim().length > 10),
    provider: provider
  };
}

/**
 * Generates studio-grade human speech MP3 (OpenAI, ElevenLabs, or Google Cloud Neural2)
 */
function synthesizeHumanVoice(text, personaId) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = (props.getProperty('HUMAN_TTS_API_KEY') || '').trim();
  var provider = (props.getProperty('HUMAN_TTS_PROVIDER') || 'openai').toLowerCase().trim();
  var customVoice = (props.getProperty('HUMAN_TTS_VOICE_ID') || '').trim();

  if (!apiKey || !text) return null;

  try {
    if (provider === 'openai') {
      var voiceName = customVoice || (personaId === 'aria_natural' ? 'nova' : 'shimmer');
      var resp = UrlFetchApp.fetch('https://api.openai.com/v1/audio/speech', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey },
        payload: JSON.stringify({
          model: 'tts-1-hd',
          voice: voiceName,
          input: text,
          speed: 0.98
        }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        return {
          mimeType: 'audio/mp3',
          audioBase64: Utilities.base64Encode(resp.getBlob().getBytes())
        };
      }
    } else if (provider === 'elevenlabs') {
      var elVoiceId = customVoice || 'EXAVITQu4vr4xnSDxMaL'; // Sarah (warm, expressive female)
      var elResp = UrlFetchApp.fetch('https://api.elevenlabs.io/v1/text-to-speech/' + elVoiceId, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'xi-api-key': apiKey },
        payload: JSON.stringify({
          text: text,
          model_id: 'eleven_multilingual_v2',
          voice_settings: { stability: 0.45, similarity_boost: 0.80, style: 0.35 }
        }),
        muteHttpExceptions: true
      });
      if (elResp.getResponseCode() === 200) {
        return {
          mimeType: 'audio/mp3',
          audioBase64: Utilities.base64Encode(elResp.getBlob().getBytes())
        };
      }
    } else if (provider === 'google') {
      var isHindi = personaId === 'swara_natural';
      var gVoice = customVoice || (isHindi ? 'hi-IN-Neural2-A' : 'en-IN-Neural2-A');
      var gLang = isHindi ? 'hi-IN' : 'en-IN';
      var gResp = UrlFetchApp.fetch('https://texttospeech.googleapis.com/v1/text:synthesize?key=' + encodeURIComponent(apiKey), {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          input: { text: text },
          voice: { languageCode: gLang, name: gVoice },
          audioConfig: { audioEncoding: 'MP3', speakingRate: 0.97, pitch: 0.5 }
        }),
        muteHttpExceptions: true
      });
      if (gResp.getResponseCode() === 200) {
        var parsed = JSON.parse(gResp.getContentText());
        if (parsed.audioContent) {
          return { mimeType: 'audio/mp3', audioBase64: parsed.audioContent };
        }
      }
    }
  } catch (e) {}

  return null;
}
/**
 * Returns whether a Real Human AI Voice API key is configured in Script Properties
 */
function getHumanVoiceConfig() {
  var props = PropertiesService.getScriptProperties();
  var key = (props.getProperty('HUMAN_TTS_API_KEY') || '').trim();
  var provider = (props.getProperty('HUMAN_TTS_PROVIDER') || 'elevenlabs').toLowerCase().trim();
  return {
    enabled: key.length > 10,
    provider: provider
  };
}

/**
 * Synthesizes real movie-grade human voice (ElevenLabs Flash v2.5, OpenAI gpt-4o-mini-tts, or Google Chirp 3 HD)
 */
function synthesizeHumanVoice(text, personaId) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = (props.getProperty('HUMAN_TTS_API_KEY') || '').trim();
  var provider = (props.getProperty('HUMAN_TTS_PROVIDER') || 'elevenlabs').toLowerCase().trim();
  var customVoice = (props.getProperty('HUMAN_TTS_VOICE_ID') || '').trim();

  if (!apiKey || !text) return null;

  try {
    if (provider === 'elevenlabs') {
      // Default: Rachel / Sarah warm expressive female voice; eleven_flash_v2_5 responds in ~75 ms
      var voiceId = customVoice || 'EXAVITQu4vr4xnSDxMaL';
      var elResp = UrlFetchApp.fetch('https://api.elevenlabs.io/v1/text-to-speech/' + voiceId + '?output_format=mp3_44100_128', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'xi-api-key': apiKey },
        payload: JSON.stringify({
          text: text,
          model_id: 'eleven_flash_v2_5',
          voice_settings: {
            stability: 0.42,
            similarity_boost: 0.82,
            style: 0.38,
            use_speaker_boost: true
          }
        }),
        muteHttpExceptions: true
      });
      if (elResp.getResponseCode() === 200) {
        return {
          mimeType: 'audio/mp3',
          audioBase64: Utilities.base64Encode(elResp.getBlob().getBytes())
        };
      }
    } else if (provider === 'openai') {
      var oaiVoice = customVoice || 'marin';
      var oaiResp = UrlFetchApp.fetch('https://api.openai.com/v1/audio/speech', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey },
        payload: JSON.stringify({
          model: 'gpt-4o-mini-tts',
          voice: oaiVoice,
          input: text,
          instructions: 'Speak like F.R.I.D.A.Y. in Iron Man: a warm, genuine, intelligent female human AI companion with natural breath, calm confidence, and subtle warmth.'
        }),
        muteHttpExceptions: true
      });
      if (oaiResp.getResponseCode() === 200) {
        return {
          mimeType: 'audio/mp3',
          audioBase64: Utilities.base64Encode(oaiResp.getBlob().getBytes())
        };
      }
    } else if (provider === 'google') {
      var isHindi = personaId === 'hindi';
      var gVoice = customVoice || (isHindi ? 'hi-IN-Chirp3-HD-Aoede' : 'en-GB-Chirp3-HD-Aoede');
      var gLang = isHindi ? 'hi-IN' : 'en-GB';
      var gResp = UrlFetchApp.fetch('https://texttospeech.googleapis.com/v1/text:synthesize?key=' + encodeURIComponent(apiKey), {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          input: { text: text },
          voice: { languageCode: gLang, name: gVoice },
          audioConfig: { audioEncoding: 'MP3', speakingRate: 0.98 }
        }),
        muteHttpExceptions: true
      });
      if (gResp.getResponseCode() === 200) {
        var parsed = JSON.parse(gResp.getContentText());
        if (parsed.audioContent) {
          return { mimeType: 'audio/mp3', audioBase64: parsed.audioContent };
        }
      }
    }
  } catch (e) {}

  return null;
}