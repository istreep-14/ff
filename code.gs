/**
 * NFL Projections Aggregator for Google Apps Script (Spreadsheet-bound)
 *
 * Features
 * - Custom menu: Setup Sheets, Refresh Data, Compute Rankings, Run All
 * - Sources sheet: define multiple sources (CSV/JSON/Sheet) with standard headers
 * - Aggregation: normalizes headers, averages projections across sources per player
 * - Settings sheet: league and scoring settings (teams, starters, roster, scoring)
 * - Output: Players sheet with points, ranks, VOR and VOLS
 *
 * Notes
 * - For non-standard source headers, add synonyms below or rename columns in the source.
 * - Avoid scraping gated sites. Prefer public CSV/JSON or your own Sheets.
 */

/**
 * Adds custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NFL Projections')
    .addItem('Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('Refresh Data (fetch + aggregate)', 'refreshData')
    .addItem('Compute Rankings (points + VOR/VOLS)', 'computeRankings')
    .addSeparator()
    .addItem('Import Sleeper Players', 'importSleeperPlayers')
    .addItem('Import DynastyProcess IDs', 'importDynastyProcessPlayerIds')
    .addSeparator()
    .addItem('Run All', 'runAll')
    .addToUi();
}

/**
 * Entry point: sets up sheets, refreshes data and computes rankings.
 */
function runAll() {
  setupSheets();
  refreshData();
  computeRankings();
}

/**
 * Creates sheets with headers if missing. Clears existing data (not settings).
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActive();

  // Settings
  const settingsSheet = getOrCreateSheet(ss, 'Settings');
  if (settingsSheet.getLastRow() <= 1) {
    const settingsHeaders = [
      'Key',
      'Value',
      'Notes'
    ];
    settingsSheet.clear();
    settingsSheet.getRange(1, 1, 1, settingsHeaders.length).setValues([settingsHeaders]);
    const defaultSettings = getDefaultSettingsRows();
    settingsSheet.getRange(2, 1, defaultSettings.length, settingsHeaders.length).setValues(defaultSettings);
    autoResizeColumns(settingsSheet, settingsHeaders.length);
  }

  // Sources
  const sourcesSheet = getOrCreateSheet(ss, 'Sources');
  const sourcesHeaders = [
    'Enabled',
    'SourceName',
    'Type',
    'URL_or_SpreadsheetID',
    'Range_or_JSONPath',
    'Notes'
  ];
  if (sourcesSheet.getLastRow() <= 1) {
    sourcesSheet.clear();
    sourcesSheet.getRange(1, 1, 1, sourcesHeaders.length).setValues([sourcesHeaders]);
    sourcesSheet.getRange(2, 1, 1, sourcesHeaders.length).setValues([[true, 'Example CSV', 'CSV', 'https://example.com/projections.csv', '', 'Headers must be standard.']]);
    sourcesSheet.getRange(3, 1, 1, sourcesHeaders.length).setValues([[false, 'Example JSON', 'JSON', 'https://example.com/projections.json', '$.data', 'Array path optional']]);
    sourcesSheet.getRange(4, 1, 1, sourcesHeaders.length).setValues([[false, 'My Sheet', 'SHEET', 'SPREADSHEET_ID_HERE', 'Sheet1!A1:Z', 'Must have standard headers']]);
    autoResizeColumns(sourcesSheet, sourcesHeaders.length);
  }

  // Players (output)
  const playersSheet = getOrCreateSheet(ss, 'Players');
  const playerHeaders = getPlayersOutputHeaders();
  playersSheet.clear();
  playersSheet.getRange(1, 1, 1, playerHeaders.length).setValues([playerHeaders]);
  autoResizeColumns(playersSheet, playerHeaders.length);

  // Logs
  const logsSheet = getOrCreateSheet(ss, 'Logs');
  if (logsSheet.getLastRow() === 0) {
    logsSheet.getRange(1, 1, 1, 3).setValues([[
      'Timestamp', 'Level', 'Message'
    ]]);
    autoResizeColumns(logsSheet, 3);
  }
}

/**
 * Fetches sources, normalizes and aggregates projections, writes to Players sheet (without rankings).
 */
function refreshData() {
  const ss = SpreadsheetApp.getActive();
  const sources = readSources();
  if (sources.length === 0) {
    logInfo('No sources enabled. Enable at least one in the Sources sheet.');
    return;
  }

  const datasets = [];
  for (var i = 0; i < sources.length; i++) {
    const source = sources[i];
    try {
      const rows = fetchSourceRows(source);
      const normalized = normalizeDataset(rows, source.SourceName);
      logInfo('Fetched ' + normalized.length + ' rows from ' + source.SourceName);
      datasets.push(normalized);
    } catch (error) {
      logError('Failed source ' + source.SourceName + ': ' + (error && error.message ? error.message : error));
    }
  }

  if (datasets.length === 0) {
    logError('No datasets fetched.');
    return;
  }

  const aggregated = aggregateDatasets(datasets);
  writePlayers(aggregated, /*includeComputedColumns*/ false);
}

/**
 * Import: Sleeper all NFL players into a sheet "Sleeper_Players".
 */
function importSleeperPlayers() {
  const ss = SpreadsheetApp.getActive();
  const sheet = getOrCreateSheet(ss, 'Sleeper_Players');
  const url = 'https://api.sleeper.app/v1/players/nfl';
  logInfo('Fetching Sleeper players...');
  const jsonText = httpGetText_(url);
  const obj = JSON.parse(jsonText);
  // Sleeper returns an object keyed by player_id → flatten to rows
  const rows = [];
  const headers = ['player_id','full_name','first_name','last_name','position','team','birth_date','height','weight','age','status','years_exp','college','active'];
  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    var p = obj[key] || {};
    rows.push([
      String(p.player_id || key || ''),
      String(p.full_name || ''),
      String(p.first_name || ''),
      String(p.last_name || ''),
      String(p.position || ''),
      String(p.team || ''),
      String(p.birth_date || ''),
      String(p.height || ''),
      String(p.weight || ''),
      p.age || '',
      String(p.status || ''),
      p.years_exp || '',
      String(p.college || ''),
      (p.active === true ? true : (p.active === false ? false : ''))
    ]);
  }
  sheet.clear();
  if (rows.length === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    autoResizeColumns(sheet, headers.length);
  }
  logInfo('Sleeper players imported: ' + rows.length);
}

/**
 * Import: DynastyProcess player ID map into sheet "DP_PlayerIDs".
 * Source: https://github.com/dynastyprocess/data/tree/master/files
 */
function importDynastyProcessPlayerIds() {
  const ss = SpreadsheetApp.getActive();
  const sheet = getOrCreateSheet(ss, 'DP_PlayerIDs');
  const csvUrl = 'https://raw.githubusercontent.com/dynastyprocess/data/master/files/playerids.csv';
  logInfo('Fetching DynastyProcess player IDs...');
  const csvText = httpGetText_(csvUrl);
  const rowsObj = parseCsvToObjects(csvText);
  // Determine headers from keys
  var headers = [];
  if (rowsObj.length > 0) {
    headers = Object.keys(rowsObj[0]);
  }
  // Convert to 2D array
  const rows = rowsObj.map(function(r){
    return headers.map(function(h){ return r[h] || ''; });
  });
  sheet.clear();
  if (headers.length > 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }
  if (rows.length > 0) {
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    autoResizeColumns(sheet, headers.length);
  }
  logInfo('DynastyProcess IDs imported: ' + rows.length);
}

/**
 * Computes points, ranks, VOR and VOLS based on current Players sheet projections and Settings.
 */
function computeRankings() {
  const ss = SpreadsheetApp.getActive();
  const playersSheet = ss.getSheetByName('Players');
  if (!playersSheet) {
    throw new Error('Players sheet not found. Run Setup Sheets first.');
  }
  const headers = playersSheet.getRange(1, 1, 1, playersSheet.getLastColumn()).getValues()[0];
  const rows = playersSheet.getRange(2, 1, Math.max(playersSheet.getLastRow() - 1, 0), headers.length).getValues();
  const players = rows
    .filter(function(r) { return r.join('').trim() !== ''; })
    .map(function(r) { return rowToPlayer(headers, r); });

  if (players.length === 0) {
    logError('No players to compute. Refresh Data first.');
    return;
  }

  const settings = readSettings();
  const scored = computePoints(players, settings.scoring);
  const ranked = computeVorVol(scored, settings);
  writePlayers(ranked, /*includeComputedColumns*/ true);
}

// ----------------------------
// Data model and helpers
// ----------------------------

/**
 * Standardized projection fields. Data sources must provide these headers (case-insensitive, synonyms supported).
 */
function getStandardFields() {
  return [
    'Player', 'Team', 'Pos', 'ByeWeek',
    'PassYds', 'PassTD', 'Int',
    'RushAtt', 'RushYds', 'RushTD',
    'Targets', 'Rec', 'RecYds', 'RecTD',
    'Fumbles', 'TwoPt', 'ReturnTD'
  ];
}

/**
 * Output headers for Players sheet.
 */
function getPlayersOutputHeaders() {
  return [
    'Player', 'Team', 'Pos', 'ByeWeek',
    'SourcesCount',
    'PassYds', 'PassTD', 'Int',
    'RushAtt', 'RushYds', 'RushTD',
    'Targets', 'Rec', 'RecYds', 'RecTD',
    'Fumbles', 'TwoPt', 'ReturnTD',
    'FantasyPoints', 'RankOverall', 'RankPos', 'VOR', 'VOLS'
  ];
}

/**
 * Default settings rows for the Settings sheet.
 */
function getDefaultSettingsRows() {
  return [
    ['NumTeams', 12, 'Number of teams in the league'],
    ['RosterSize', 16, 'Total roster size per team'],
    ['Starters_QB', 1, 'Starting QBs per team'],
    ['Starters_RB', 2, 'Starting RBs per team'],
    ['Starters_WR', 2, 'Starting WRs per team'],
    ['Starters_TE', 1, 'Starting TEs per team'],
    ['Starters_FLEX', 1, 'Flex spots (RB/WR/TE)'],
    ['Starters_K', 0, 'Kickers (not computed unless provided)'],
    ['Starters_DST', 0, 'Defenses (not computed)'],
    ['Bench_QB', 0, 'Bench QBs per team for VOR baseline'],
    ['Bench_RB', 1, 'Bench RBs per team for VOR baseline'],
    ['Bench_WR', 1, 'Bench WRs per team for VOR baseline'],
    ['Bench_TE', 0, 'Bench TEs per team for VOR baseline'],
    ['PPR', 1, '1=PPR, 0=Standard, 0.5=Half PPR'],
    ['PassYdPerPoint', 25, 'Passing yards per fantasy point'],
    ['PassTDPts', 4, 'Points per passing TD'],
    ['IntPts', -2, 'Points per interception'],
    ['RushYdPerPoint', 10, 'Rushing yards per fantasy point'],
    ['RushTDPts', 6, 'Points per rushing TD'],
    ['RecPt', 1, 'Points per reception (redundant with PPR for flexibility)'],
    ['RecYdPerPoint', 10, 'Receiving yards per fantasy point'],
    ['RecTDPts', 6, 'Points per receiving TD'],
    ['FumblePts', -2, 'Points per fumble lost'],
    ['TwoPtPts', 2, 'Points per 2-point conversion'],
    ['ReturnTDPts', 6, 'Points per return TD'],
    ['UseCustomVorRanks', 0, '1=use VOR_Rank_* if provided'],
    ['VOR_Rank_QB', '', 'Optional custom replacement rank for QB'],
    ['VOR_Rank_RB', '', 'Optional custom replacement rank for RB'],
    ['VOR_Rank_WR', '', 'Optional custom replacement rank for WR'],
    ['VOR_Rank_TE', '', 'Optional custom replacement rank for TE'],
    ['VOL_UseAuto', 1, '1=auto last starter ranks; 0=use VOL_Rank_*'],
    ['VOL_Rank_QB', '', 'Optional custom last starter rank for QB'],
    ['VOL_Rank_RB', '', 'Optional custom last starter rank for RB'],
    ['VOL_Rank_WR', '', 'Optional custom last starter rank for WR'],
    ['VOL_Rank_TE', '', 'Optional custom last starter rank for TE']
  ];
}

/**
 * Reads settings from the Settings sheet.
 */
function readSettings() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    throw new Error('Settings sheet not found. Run Setup Sheets first.');
  }
  const rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 3).getValues();
  const map = {};
  for (var i = 0; i < rows.length; i++) {
    const key = String(rows[i][0] || '').trim();
    if (!key) continue;
    map[key] = rows[i][1];
  }

  const scoring = {
    ppr: num(map['PPR'], 1),
    passYdPerPoint: num(map['PassYdPerPoint'], 25),
    passTDPts: num(map['PassTDPts'], 4),
    intPts: num(map['IntPts'], -2),
    rushYdPerPoint: num(map['RushYdPerPoint'], 10),
    rushTDPts: num(map['RushTDPts'], 6),
    recPt: num(map['RecPt'], 1),
    recYdPerPoint: num(map['RecYdPerPoint'], 10),
    recTDPts: num(map['RecTDPts'], 6),
    fumblePts: num(map['FumblePts'], -2),
    twoPtPts: num(map['TwoPtPts'], 2),
    returnTDPts: num(map['ReturnTDPts'], 6)
  };

  const league = {
    numTeams: num(map['NumTeams'], 12),
    rosterSize: num(map['RosterSize'], 16),
    starters: {
      QB: num(map['Starters_QB'], 1),
      RB: num(map['Starters_RB'], 2),
      WR: num(map['Starters_WR'], 2),
      TE: num(map['Starters_TE'], 1),
      FLEX: num(map['Starters_FLEX'], 1),
      K: num(map['Starters_K'], 0),
      DST: num(map['Starters_DST'], 0)
    },
    bench: {
      QB: num(map['Bench_QB'], 0),
      RB: num(map['Bench_RB'], 1),
      WR: num(map['Bench_WR'], 1),
      TE: num(map['Bench_TE'], 0)
    }
  };

  const vorCustomRanksEnabled = num(map['UseCustomVorRanks'], 0) === 1;
  const volAuto = num(map['VOL_UseAuto'], 1) === 1;

  const vorRanks = {
    QB: map['VOR_Rank_QB'],
    RB: map['VOR_Rank_RB'],
    WR: map['VOR_Rank_WR'],
    TE: map['VOR_Rank_TE']
  };
  const volRanks = {
    QB: map['VOL_Rank_QB'],
    RB: map['VOL_Rank_RB'],
    WR: map['VOL_Rank_WR'],
    TE: map['VOL_Rank_TE']
  };

  return {
    league: league,
    scoring: scoring,
    vorCustomRanksEnabled: vorCustomRanksEnabled,
    vorRanks: vorRanks,
    volAuto: volAuto,
    volRanks: volRanks
  };
}

// ----------------------------
// Source ingestion
// ----------------------------

/**
 * Reads enabled sources from Sources sheet.
 */
function readSources() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Sources');
  if (!sheet) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).getValues();
  const idx = indexMap(headers);
  const out = [];
  for (var i = 0; i < rows.length; i++) {
    const r = rows[i];
    const enabled = asBool(r[idx['Enabled']]);
    if (!enabled) continue;
    const source = {
      Enabled: enabled,
      SourceName: String(r[idx['SourceName']] || '').trim() || 'Source_' + (i + 1),
      Type: String(r[idx['Type']] || '').trim().toUpperCase(),
      URL_or_SpreadsheetID: String(r[idx['URL_or_SpreadsheetID']] || '').trim(),
      Range_or_JSONPath: String(r[idx['Range_or_JSONPath']] || '').trim()
    };
    if (!source.Type || !source.URL_or_SpreadsheetID) continue;
    out.push(source);
  }
  return out;
}

/**
 * Fetches rows for a source. Returns array of objects with at least standard field keys.
 */
function fetchSourceRows(source) {
  if (source.Type === 'CSV') {
    const csv = httpGetText_(source.URL_or_SpreadsheetID);
    return parseCsvToObjects(csv);
  }
  if (source.Type === 'JSON') {
    const text = httpGetText_(source.URL_or_SpreadsheetID);
    var json = JSON.parse(text);
    if (source.Range_or_JSONPath && source.Range_or_JSONPath.trim() !== '') {
      json = jsonPath(json, source.Range_or_JSONPath);
    }
    if (!Array.isArray(json)) {
      throw new Error('JSON path did not resolve to an array for ' + source.SourceName);
    }
    return json.map(function(o) { return o; });
  }
  if (source.Type === 'SHEET') {
    const range = source.Range_or_JSONPath || 'Sheet1!A1:Z';
    const ss = SpreadsheetApp.openById(source.URL_or_SpreadsheetID);
    const a1 = range;
    const rng = ss.getRange(a1);
    const values = rng.getValues();
    const headers = (values[0] || []).map(function(h) { return String(h || '').trim(); });
    const out = [];
    for (var i = 1; i < values.length; i++) {
      const row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = values[i][j];
      }
      if (Object.keys(row).length > 0) out.push(row);
    }
    return out;
  }
  throw new Error('Unsupported source type: ' + source.Type);
}

/**
 * Normalizes a dataset's headers to the standard field names.
 */
function normalizeDataset(rows, sourceName) {
  const standard = getStandardFields();
  const synonyms = getHeaderSynonyms();
  const out = [];
  for (var i = 0; i < rows.length; i++) {
    const r = rows[i];
    const o = {};
    for (var k = 0; k < standard.length; k++) {
      var field = standard[k];
      var value = r[field];
      if (value === undefined) {
        // try synonyms
        const synList = synonyms[field] || [];
        for (var s = 0; s < synList.length; s++) {
          const alt = synList[s];
          if (r.hasOwnProperty(alt)) { value = r[alt]; break; }
        }
      }
      o[field] = value;
    }
    // Required fields: Player, Pos
    const playerName = String(o['Player'] || '').trim();
    const pos = String(o['Pos'] || '').trim().toUpperCase();
    if (!playerName || !pos) continue;
    // Normalize numbers
    numericNormalize(o, ['PassYds','PassTD','Int','RushAtt','RushYds','RushTD','Targets','Rec','RecYds','RecTD','Fumbles','TwoPt','ReturnTD']);
    o['Team'] = String(o['Team'] || '').toUpperCase();
    o['Pos'] = pos;
    o['ByeWeek'] = num(o['ByeWeek'], '');
    out.push(o);
  }
  return out;
}

/**
 * Aggregates multiple datasets by player (Player + Pos + Team if present). Averages numeric fields; counts sources.
 */
function aggregateDatasets(datasets) {
  const keyFn = function(p) {
    const nameKey = normalizeName(p.Player);
    const teamKey = (p.Team && String(p.Team).trim() !== '') ? String(p.Team).toUpperCase() : 'NA';
    return nameKey + '|' + p.Pos + '|' + teamKey;
  };
  const fieldList = ['PassYds','PassTD','Int','RushAtt','RushYds','RushTD','Targets','Rec','RecYds','RecTD','Fumbles','TwoPt','ReturnTD'];
  const map = {};
  for (var d = 0; d < datasets.length; d++) {
    const rows = datasets[d];
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      const key = keyFn(r);
      if (!map[key]) {
        map[key] = {
          Player: r.Player,
          Team: r.Team || '',
          Pos: r.Pos,
          ByeWeek: r.ByeWeek || '',
          SourcesCount: 0
        };
        for (var f = 0; f < fieldList.length; f++) map[key][fieldList[f]] = 0;
      }
      map[key].SourcesCount++;
      for (var f2 = 0; f2 < fieldList.length; f2++) {
        const fld = fieldList[f2];
        map[key][fld] += num(r[fld], 0);
      }
    }
  }
  const out = [];
  Object.keys(map).forEach(function(k) {
    const o = map[k];
    if (o.SourcesCount > 0) {
      for (var f = 0; f < fieldList.length; f++) {
        const fld = fieldList[f];
        o[fld] = round(o[fld] / o.SourcesCount, 3);
      }
    }
    out.push(o);
  });
  return out;
}

// ----------------------------
// Scoring and VOR/VOLS
// ----------------------------

/**
 * Computes fantasy points for each player given scoring settings.
 */
function computePoints(players, scoring) {
  return players.map(function(p) {
    const points =
      (num(p.PassYds, 0) / num(scoring.passYdPerPoint, 25)) +
      (num(p.PassTD, 0) * num(scoring.passTDPts, 4)) +
      (num(p.Int, 0) * num(scoring.intPts, -2)) +
      (num(p.RushYds, 0) / num(scoring.rushYdPerPoint, 10)) +
      (num(p.RushTD, 0) * num(scoring.rushTDPts, 6)) +
      (num(p.Rec, 0) * (num(scoring.recPt, 1) || num(scoring.ppr, 1))) +
      (num(p.RecYds, 0) / num(scoring.recYdPerPoint, 10)) +
      (num(p.RecTD, 0) * num(scoring.recTDPts, 6)) +
      (num(p.Fumbles, 0) * num(scoring.fumblePts, -2)) +
      (num(p.TwoPt, 0) * num(scoring.twoPtPts, 2)) +
      (num(p.ReturnTD, 0) * num(scoring.returnTDPts, 6));
    const out = clone(p);
    out.FantasyPoints = round(points, 3);
    return out;
  });
}

/**
 * Computes ranks, VOR and VOLS.
 */
function computeVorVol(players, settings) {
  // Split by position
  const byPos = groupBy(players, function(p) { return p.Pos; });
  const positions = ['QB','RB','WR','TE'];

  // Rank within position
  const rankedByPos = {};
  positions.forEach(function(pos) {
    const list = (byPos[pos] || []).slice().sort(function(a, b) { return b.FantasyPoints - a.FantasyPoints; });
    for (var i = 0; i < list.length; i++) list[i].RankPos = i + 1;
    rankedByPos[pos] = list;
  });

  // Overall ranks
  const overall = players.slice().sort(function(a, b) { return b.FantasyPoints - a.FantasyPoints; });
  for (var i = 0; i < overall.length; i++) overall[i].RankOverall = i + 1;

  // Determine baseline ranks for VOR (replacement) and VOL (last starter)
  const numTeams = settings.league.numTeams;
  const starters = settings.league.starters;
  const bench = settings.league.bench;

  const volRanks = {};
  const vorRanks = {};

  positions.forEach(function(pos) {
    // VOL baseline: last starter per position plus flex allocation if auto, else custom rank
    if (settings.volAuto) {
      var posStarters = starters[pos] || 0;
      if (pos === 'RB' || pos === 'WR' || pos === 'TE') {
        // Allocate FLEX across RB/WR/TE using equal split
        const flexSlots = starters.FLEX || 0;
        const perPosFlex = Math.floor(flexSlots / 3);
        posStarters += perPosFlex;
      }
      volRanks[pos] = numTeams * Math.max(1, posStarters);
    } else {
      volRanks[pos] = num(settings.volRanks[pos], '') || (numTeams * Math.max(1, starters[pos] || 0));
    }

    // VOR baseline: starters + bench per team for the position (replacement level)
    const base = (starters[pos] || 0) + (bench[pos] || 0);
    vorRanks[pos] = settings.vorCustomRanksEnabled && num(settings.vorRanks[pos], '')
      ? num(settings.vorRanks[pos], '')
      : numTeams * Math.max(1, base);
  });

  // Get baseline points for each position
  const baselinePointsVOL = {};
  const baselinePointsVOR = {};
  positions.forEach(function(pos) {
    const list = rankedByPos[pos] || [];
    baselinePointsVOL[pos] = (list[clamp(volRanks[pos] - 1, 0, list.length - 1)] || {FantasyPoints: 0}).FantasyPoints || 0;
    baselinePointsVOR[pos] = (list[clamp(vorRanks[pos] - 1, 0, list.length - 1)] || {FantasyPoints: 0}).FantasyPoints || 0;
  });

  // Compute VOR and VOLS
  const out = players.map(function(p) {
    const vor = p.FantasyPoints - (baselinePointsVOR[p.Pos] || 0);
    const vols = p.FantasyPoints - (baselinePointsVOL[p.Pos] || 0);
    const o = clone(p);
    o.VOR = round(vor, 3);
    o.VOLS = round(vols, 3);
    return o;
  });

  // Recompute overall and position ranks on output to ensure fields exist
  const outByPos = groupBy(out, function(p) { return p.Pos; });
  positions.forEach(function(pos) {
    const list = (outByPos[pos] || []).slice().sort(function(a, b) { return b.FantasyPoints - a.FantasyPoints; });
    for (var i = 0; i < list.length; i++) list[i].RankPos = i + 1;
  });
  const outOverall = out.slice().sort(function(a, b) { return b.FantasyPoints - a.FantasyPoints; });
  for (var j = 0; j < outOverall.length; j++) outOverall[j].RankOverall = j + 1;

  return out;
}

// ----------------------------
// Sheet I/O
// ----------------------------

function writePlayers(players, includeComputedColumns) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Players');
  if (!sheet) throw new Error('Players sheet not found.');
  const headers = getPlayersOutputHeaders();
  const rows = [];
  for (var i = 0; i < players.length; i++) {
    const p = players[i];
    const row = [
      p.Player || '', p.Team || '', p.Pos || '', p.ByeWeek || '',
      p.SourcesCount || 1,
      num(p.PassYds, 0), num(p.PassTD, 0), num(p.Int, 0),
      num(p.RushAtt, 0), num(p.RushYds, 0), num(p.RushTD, 0),
      num(p.Targets, 0), num(p.Rec, 0), num(p.RecYds, 0), num(p.RecTD, 0),
      num(p.Fumbles, 0), num(p.TwoPt, 0), num(p.ReturnTD, 0)
    ];
    if (includeComputedColumns) {
      row.push(num(p.FantasyPoints, 0));
      row.push(num(p.RankOverall, ''));
      row.push(num(p.RankPos, ''));
      row.push(num(p.VOR, 0));
      row.push(num(p.VOLS, 0));
    } else {
      row.push(''); // FantasyPoints
      row.push(''); // RankOverall
      row.push(''); // RankPos
      row.push(''); // VOR
      row.push(''); // VOLS
    }
    rows.push(row);
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  autoResizeColumns(sheet, headers.length);
}

function rowToPlayer(headers, row) {
  const idx = indexMap(headers);
  return {
    Player: row[idx['Player']],
    Team: row[idx['Team']],
    Pos: row[idx['Pos']],
    ByeWeek: row[idx['ByeWeek']],
    SourcesCount: num(row[idx['SourcesCount']], 1),
    PassYds: num(row[idx['PassYds']], 0),
    PassTD: num(row[idx['PassTD']], 0),
    Int: num(row[idx['Int']], 0),
    RushAtt: num(row[idx['RushAtt']], 0),
    RushYds: num(row[idx['RushYds']], 0),
    RushTD: num(row[idx['RushTD']], 0),
    Targets: num(row[idx['Targets']], 0),
    Rec: num(row[idx['Rec']], 0),
    RecYds: num(row[idx['RecYds']], 0),
    RecTD: num(row[idx['RecTD']], 0),
    Fumbles: num(row[idx['Fumbles']], 0),
    TwoPt: num(row[idx['TwoPt']], 0),
    ReturnTD: num(row[idx['ReturnTD']], 0),
    FantasyPoints: num(row[idx['FantasyPoints']], 0),
    RankOverall: num(row[idx['RankOverall']], ''),
    RankPos: num(row[idx['RankPos']], ''),
    VOR: num(row[idx['VOR']], 0),
    VOLS: num(row[idx['VOLS']], 0)
  };
}

// ----------------------------
// Utilities
// ----------------------------

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function autoResizeColumns(sheet, count) {
  try { sheet.autoResizeColumns(1, count); } catch (e) {}
}

function indexMap(headers) {
  const map = {};
  for (var i = 0; i < headers.length; i++) map[headers[i]] = i;
  return map;
}

function num(v, def) {
  if (v === '' || v === null || v === undefined) return def;
  if (typeof v === 'number') return v;
  var n = Number(v);
  return isNaN(n) ? def : n;
}

function asBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const s = String(v || '').toLowerCase();
  return s === 'true' || s === 'yes' || s === '1' || s === 'y';
}

function round(n, digits) {
  var m = Math.pow(10, digits || 0);
  return Math.round(n * m) / m;
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function clone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function groupBy(arr, keyFn) {
  const m = {};
  for (var i = 0; i < arr.length; i++) {
    const key = keyFn(arr[i]);
    if (!m[key]) m[key] = [];
    m[key].push(arr[i]);
  }
  return m;
}

function normalizeName(name) {
  return String(name || '').trim().toUpperCase().replace(/[^A-Z ]+/g, '');
}

function numericNormalize(obj, fields) {
  for (var i = 0; i < fields.length; i++) {
    var f = fields[i];
    obj[f] = num(obj[f], 0);
  }
}

function logInfo(message) {
  log_('INFO', message);
}

function logError(message) {
  log_('ERROR', message);
}

function log_(level, message) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Logs');
    if (!sheet) return;
    sheet.appendRow([new Date(), level, String(message)]);
  } catch (e) {
    // no-op
  }
}

// ----------------------------
// Header synonyms and parsers
// ----------------------------

function getHeaderSynonyms() {
  return {
    Player: ['Name','player','PLAYER'],
    Team: ['Tm','team','TEAM','NFLTeam'],
    Pos: ['Position','POS','position'],
    ByeWeek: ['Bye','BYE','bye'],
    PassYds: ['PassYards','passYds','PaYds','pass_yds','PY'],
    PassTD: ['PassTouchdowns','PaTD','pass_td','PassTds'],
    Int: ['Interceptions','INT','Ints'],
    RushAtt: ['RushAttempts','Att','rush_att'],
    RushYds: ['RushYards','RuYds','rush_yds'],
    RushTD: ['RushTouchdowns','RuTD','rush_td'],
    Targets: ['Tgt','targets'],
    Rec: ['Receptions','REC','receptions'],
    RecYds: ['ReceivingYards','ReYds','rec_yds'],
    RecTD: ['ReceivingTouchdowns','ReTD','rec_td'],
    Fumbles: ['FumblesLost','FL','fum'],
    TwoPt: ['TwoPoint','TwoPointConv','2PT','two_pt'],
    ReturnTD: ['RetTD','ret_td']
  };
}

function parseCsvToObjects(csvText) {
  const data = Utilities.parseCsv(csvText);
  if (!data || data.length === 0) return [];
  const headers = data[0].map(function(h) { return String(h || '').trim(); });
  const out = [];
  for (var i = 1; i < data.length; i++) {
    const row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    if (Object.keys(row).length > 0) out.push(row);
  }
  return out;
}

function httpGetText_(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' fetching ' + url);
  }
  return resp.getContentText();
}

/**
 * Simple JSONPath: supports top-level '$' and one level of property selection like '$.data.items'
 */
function jsonPath(obj, path) {
  var p = String(path || '').trim();
  if (p === '' || p === '$') return obj;
  if (p.indexOf('$.') === 0) p = p.substring(2);
  var parts = p.split('.');
  var cur = obj;
  for (var i = 0; i < parts.length; i++) {
    if (cur == null) return null;
    cur = cur[parts[i]];
  }
  return cur;
}
