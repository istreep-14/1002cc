// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  USERNAME: 'ians141',
  MAX_GAMES_PER_BATCH: 3,
  AUTO_FETCH_CALLBACK_DATA: true, 
  AUTO_REFRESH_DAILY_STATS: true  
};

const SHEETS = {
  GAMES: 'Games',
  CALLBACK: 'Callback',
  DERIVED: 'Derived Data'
};


// ============================================
// CALLBACK DATA FETCHING
// ============================================


function fetchCallbackData(game) {
  // Validate game object has required fields
  if (!game || !game.gameId || !game.timeClass || !game.white || !game.black) {
    Logger.log(`Skipping callback fetch - incomplete game data: ${JSON.stringify(game)}`);
    return null;
  }
  
  const gameId = game.gameId;
  const timeClass = game.time_class || 'unknown';
  const gameType = timeClass === 'daily' ? 'daily' : 'live';
  const callbackUrl = `https://www.chess.com/callback/${gameType}/game/${gameId}`;
  
  try {
    const response = UrlFetchApp.fetch(callbackUrl, {muteHttpExceptions: true});
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Callback API error for game ${gameId}: ${response.getResponseCode()}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || !data.game) {
      Logger.log(`Invalid callback data for game ${gameId}`);
      return null;
    }
    
    const gameData = data.game;
    const players = data.players || {};
    const topPlayer = players.top || {};
    const bottomPlayer = players.bottom || {};
    
    // Determine my color and player data
    const isWhite = game.white === CONFIG.USERNAME;
    const myColor = isWhite ? 'white' : 'black';
    
    let myRatingChange = isWhite ? gameData.ratingChangeWhite : gameData.ratingChangeBlack;
    let oppRatingChange = isWhite ? gameData.ratingChangeBlack : gameData.ratingChangeWhite;
    
    // If rating change is 0, it's likely an error (unless edge case draw)
    // Set to null to indicate unreliable data
    if (myRatingChange === 0) myRatingChange = null;
    if (oppRatingChange === 0) oppRatingChange = null;
    
    // Get player data (top/bottom can be either color)
    let whitePlayer, blackPlayer;
    if (topPlayer.color === 'white') {
      whitePlayer = topPlayer;
      blackPlayer = bottomPlayer;
    } else {
      whitePlayer = bottomPlayer;
      blackPlayer = topPlayer;
    }
    
    // Determine my player and opponent player
    const myPlayer = isWhite ? whitePlayer : blackPlayer;
    const oppPlayer = isWhite ? blackPlayer : whitePlayer;
    
    // Get ratings from callback
    const myRating = myPlayer.rating || null;
    const oppRating = oppPlayer.rating || null;
    
    // Calculate "before" ratings by subtracting rating change
    let myRatingBefore = null;
    let oppRatingBefore = null;
    
    if (myRating !== null && myRatingChange !== null) {
      myRatingBefore = myRating - myRatingChange;
    }
    if (oppRating !== null && oppRatingChange !== null) {
      oppRatingBefore = oppRating - oppRatingChange;
    }
    
    return {
      gameId: gameId,
      gameUrl: game.gameUrl,
      callbackUrl: callbackUrl,
      endTime: gameData.endTime,
      myColor: myColor,
      timeClass: game.timeClass,
      myRating: myRating,
      oppRating: oppRating,
      myRatingChange: myRatingChange,
      oppRatingChange: oppRatingChange,
      myRatingBefore: myRatingBefore,
      oppRatingBefore: oppRatingBefore,
      baseTime: gameData.baseTime1 || 0,
      timeIncrement: gameData.timeIncrement1 || 0,
      moveTimestamps: gameData.moveTimestamps ? String(gameData.moveTimestamps) : '',
      myUsername: myPlayer.username || '',
      myCountry: myPlayer.countryName || '',
      myMembership: myPlayer.membershipCode || '',
      myMemberSince: myPlayer.memberSince || 0,
      myDefaultTab: myPlayer.defaultTab || null,
      myPostMoveAction: myPlayer.postMoveAction || '',
      myLocation: myPlayer.location || '',
      oppUsername: oppPlayer.username || '',
      oppCountry: oppPlayer.countryName || '',
      oppMembership: oppPlayer.membershipCode || '',
      oppMemberSince: oppPlayer.memberSince || 0,
      oppDefaultTab: oppPlayer.defaultTab || null,
      oppPostMoveAction: oppPlayer.postMoveAction || '',
      oppLocation: oppPlayer.location || ''
    };
    
  } catch (error) {
    Logger.log(`Error fetching callback data for game ${gameId}: ${error.message}`);
    return null;
  }
}

function saveCallbackData(callbackData) {
  if (!callbackData) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!callbackSheet) return;
  
  const row = [
    callbackData.gameId,
    callbackData.gameUrl,
    callbackData.callbackUrl,
    callbackData.endTime,
    callbackData.myColor,
    callbackData.timeClass,
    callbackData.myRating,
    callbackData.oppRating,
    callbackData.myRatingChange,
    callbackData.oppRatingChange,
    callbackData.myRatingBefore,
    callbackData.oppRatingBefore,
    callbackData.baseTime,
    callbackData.timeIncrement,
    callbackData.moveTimestamps,
    callbackData.myUsername,
    callbackData.myCountry,
    callbackData.myMembership,
    callbackData.myMemberSince,
    callbackData.myDefaultTab,
    callbackData.myPostMoveAction,
    callbackData.myLocation,
    callbackData.oppUsername,
    callbackData.oppCountry,
    callbackData.oppMembership,
    callbackData.oppMemberSince,
    callbackData.oppDefaultTab,
    callbackData.oppPostMoveAction,
    callbackData.oppLocation,
    new Date()
  ];
  
  const lastRow = callbackSheet.getLastRow();
  callbackSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function processNewGamesAutoFeatures(newGames) {
  if (!newGames || newGames.length === 0) return;
  
  // Auto-fetch callback data
  if (CONFIG.AUTO_FETCH_CALLBACK_DATA && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    fetchCallbackForGames(newGames);
  }
  
  // Auto-analyze new games
  if (CONFIG.AUTO_ANALYZE_NEW_GAMES && newGames.length <= CONFIG.MAX_GAMES_PER_BATCH) {
    analyzeGames(newGames);
  }
}

// ============================================
// FETCH CALLBACK DATA FOR GAMES
// ============================================
function fetchCallbackLast10() { fetchCallbackLastN(10); }
function fetchCallbackLast25() { fetchCallbackLastN(25); }
function fetchCallbackLast50() { fetchCallbackLastN(50); }

function fetchCallbackLastN(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  
  if (!gamesSheet || !callbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const gamesWithoutCallback = getGamesWithoutCallback(count);
  
  if (gamesWithoutCallback.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No games need callback data!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Fetch callback data for ${gamesWithoutCallback.length} game(s)?`,
    `This will fetch detailed game data from Chess.com.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  fetchCallbackForGames(gamesWithoutCallback);
}

function getGamesWithoutCallback(maxCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  const games = [];
  
  // Iterate from newest to oldest (reverse order)
  for (let i = data.length - 1; i >= 1 && games.length < maxCount; i--) {
    if (data[i][13] === false) { // Callback Fetched column (index 13)
      const myColor = data[i][3];
      const opponent = data[i][4];
      const gameId = data[i][11];
      
      // Get time class from derived data
      const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
      let timeClass = '';
      
      if (derivedSheet) {
        const derivedData = derivedSheet.getDataRange().getValues();
        for (let j = 1; j < derivedData.length; j++) {
          if (derivedData[j][0] === gameId) {
            timeClass = derivedData[j][5];
            break;
          }
        }
      }
      
      games.push({
        row: i + 1,
        gameId: gameId,
        gameUrl: data[i][0],
        white: myColor === 'white' ? CONFIG.USERNAME : opponent,
        black: myColor === 'black' ? CONFIG.USERNAME : opponent,
        timeClass: timeClass
      });
    }
  }
  
  return games.reverse(); // Return in chronological order (oldest first)
}

function fetchCallbackForGames(gamesToFetch) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  let successCount = 0;
  let errorCount = 0;
  
  ss.toast('Fetching callback data...', '📋', -1);
  
  for (let i = 0; i < gamesToFetch.length; i++) {
    const game = gamesToFetch[i];
    
    try {
      ss.toast(`Fetching callback ${i + 1} of ${gamesToFetch.length}...`, '📋', -1);
      
      const callbackData = fetchCallbackData(game);
      if (callbackData) {
        saveCallbackData(callbackData);
        
        // Mark callback as fetched
        if (game.row) {
          gamesSheet.getRange(game.row, 14).setValue(true); // Callback Fetched column (index 14)
        }
        
        successCount++;
      } else {
        errorCount++;
      }
      
      Utilities.sleep(300); // Rate limiting
      
    } catch (error) {
      Logger.log(`Error fetching callback for game ${game.gameId}: ${error}`);
      errorCount++;
    }
  }
  
  ss.toast(`✅ Callback fetched: ${successCount}, Errors: ${errorCount}`, '📋', 5);
  
}

function findGameRow(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const data = gamesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === gameId) { // Game ID column (index 11)
      return i + 1;
    }
  }
  return -1;
}

// ============================================
// FETCH CHESS.COM GAMES
// ============================================

// INITIAL FETCH: Get all games from all archives
function fetchAllGamesInitialOptimized() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Initial Full Fetch',
    'This will fetch ALL games from your Chess.com history.\n' +
    'This may take several minutes depending on how many games you have.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    ss.toast('Fetching all game archives...', '⏳', -1);
    
    const archivesUrl = `https://api.chess.com/pub/player/${username}/games/archives`;
    const archivesResponse = UrlFetchApp.fetch(archivesUrl);
    const archives = JSON.parse(archivesResponse.getContentText()).archives;
    
    ss.toast(`Found ${archives.length} archives. Fetching games...`, '⏳', -1);
    
    const allGames = [];
    const props = PropertiesService.getScriptProperties();
    const now = new Date();
    
    for (let i = 0; i < archives.length; i++) {
      ss.toast(`Fetching archive ${i + 1} of ${archives.length}...`, '⏳', -1);
      Utilities.sleep(500);
      
      const response = fetchWithETag(archives[i], null);
      if (response.data) {
        allGames.push(...response.data.games);
      }
      
      Logger.log(`Archive ${i + 1}/${archives.length}: ${response.data?.games?.length || 0} games`);
    }
    
    // Filter duplicates
    const existingGameIds = new Set();
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]);
      }
    }
    
    const newGames = allGames.filter(game => !existingGameIds.has(game.url.split('/').pop()));
    
    ss.toast(`Processing ${newGames.length} games...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Find most recent game and store its date
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      
      const lastGameDate = new Date(mostRecentGame.end_time * 1000);
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      props.setProperty('LAST_GAME_TIMESTAMP', mostRecentGame.end_time.toString());
      props.setProperty('INITIAL_FETCH_COMPLETE', 'true');
      
      ss.toast(`✅ Fetched ${newGames.length} games!`, '✅', 5);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// UPDATE FETCH: Check most recent archive(s) for new games
function fetchChesscomGamesOptimized() {
  const username = CONFIG.USERNAME;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" first!');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  const initialFetchComplete = props.getProperty('INITIAL_FETCH_COMPLETE');
  
  if (!initialFetchComplete) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'No Initial Fetch Detected',
      'It looks like you haven\'t done an initial full fetch yet.\n\n' +
      'Would you like to do a quick recent fetch (current month)?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
  }
  
  try {
    const now = new Date();
    const lastGameTimestamp = parseInt(props.getProperty('LAST_GAME_TIMESTAMP') || '0');
    const lastGameDate = lastGameTimestamp ? new Date(lastGameTimestamp * 1000) : new Date(0);
    
    // Calculate which archives to check based on last game date
    const archivesToCheck = [];
    const lastKnownGameUrl = props.getProperty('LAST_GAME_URL');
    
    // If last game is from a previous month, finalize that month first
    const lastGameYear = lastGameDate.getFullYear();
    const lastGameMonth = lastGameDate.getMonth() + 1;
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    
    if (lastGameYear < currentYear || (lastGameYear === currentYear && lastGameMonth < currentMonth)) {
      // Need to finalize the month of the last game
      const lastArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${lastGameYear}/${String(lastGameMonth).padStart(2, '0')}`;
      archivesToCheck.push({url: lastArchiveUrl, isCurrent: false});
    }
    
    // Always check current month
    const currentArchiveUrl = `https://api.chess.com/pub/player/${username}/games/${currentYear}/${String(currentMonth).padStart(2, '0')}`;
    archivesToCheck.push({url: currentArchiveUrl, isCurrent: true});
    
    const allGames = [];
    let foundLastKnownGame = false;
    
    for (const archive of archivesToCheck) {
      Utilities.sleep(500);
      
      const storedETag = archive.isCurrent ? props.getProperty('etag_current') : null;
      const response = fetchWithETag(archive.url, storedETag);
      
      if (!response.data) {
        Logger.log(`Archive ${archive.url} not modified`);
        continue;
      }
      
      if (archive.isCurrent) {
        props.setProperty('etag_current', response.etag);
      }
      
      const gamesData = response.data.games;
      
      if (lastKnownGameUrl) {
        for (let i = gamesData.length - 1; i >= 0; i--) {
          const game = gamesData[i];
          if (game.url === lastKnownGameUrl) {
            foundLastKnownGame = true;
            break;
          }
          allGames.unshift(game);
        }
        if (foundLastKnownGame) break;
      } else {
        allGames.push(...gamesData);
      }
    }
    
    if (allGames.length === 0) {
      ss.toast('No new games found!', 'ℹ️', 3);
      return;
    }
    
    // Filter duplicates
    const existingGameIds = new Set();
    if (gamesSheet.getLastRow() > 1) {
      const existingData = gamesSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        existingGameIds.add(existingData[i][11]);
      }
    }
    
    const newGames = allGames.filter(game => !existingGameIds.has(game.url.split('/').pop()));
    
    if (newGames.length === 0) {
      ss.toast('No new games found!', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Processing ${newGames.length} new games...`, '⏳', -1);
    const rows = processGamesData(newGames, username);
    
    if (rows.length > 0) {
      const lastRow = gamesSheet.getLastRow();
      gamesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      let mostRecentGame = newGames[0];
      for (const game of newGames) {
        if (game.end_time > mostRecentGame.end_time) {
          mostRecentGame = game;
        }
      }
      props.setProperty('LAST_GAME_URL', mostRecentGame.url);
      props.setProperty('LAST_GAME_TIMESTAMP', mostRecentGame.end_time.toString());
      
      ss.toast(`✅ Added ${rows.length} new games!`, '✅', 5);
      
      // Auto-refresh Daily Stats if enabled
      if (CONFIG.AUTO_REFRESH_DAILY_STATS) {
        setupDailyStatsSheet();
      }
      
      const gamesToProcess = newGames.map(g => ({
        row: findGameRow(g.url.split('/').pop()),
        gameId: g.url.split('/').pop(),
        gameUrl: g.url,
        white: g.white?.username || '',
        black: g.black?.username || '',
        outcome: getGameOutcome(g, CONFIG.USERNAME),
        pgn: g.pgn || ''
      })).filter(g => g.row > 0 && g.gameId && g.white && g.black);
      
      processNewGamesAutoFeatures(gamesToProcess);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// Fetch with ETag support
function fetchWithETag(url, etag) {
  const options = {
    muteHttpExceptions: true,
    headers: {}
  };
  
  if (etag) {
    options.headers['If-None-Match'] = etag;
  }
  
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  
  if (code === 304) {
    // Not modified
    return { data: null, etag: etag };
  }
  
  if (code === 200) {
    const newETag = response.getHeaders()['ETag'] || response.getHeaders()['etag'] || '';
    const data = JSON.parse(response.getContentText());
    return { data: data, etag: newETag };
  }
  
  throw new Error(`HTTP ${code}: ${response.getContentText()}`);
}

// Convert archive URL to storage key
function archiveUrlToKey(url) {
  // Extract YYYY/MM from URL like https://api.chess.com/pub/player/username/games/2024/09
  const match = url.match(/(\d{4})\/(\d{2})$/);
  return match ? `${match[1]}_${match[2]}` : url;
}

// Parse time control string into components
function parseTimeControl(timeControl, timeClass) {
  const result = {
    type: timeClass === 'daily' ? 'Daily' : 'Live',
    baseTime: null,
    increment: null,
    correspondenceTime: null
  };
  
  if (!timeControl) return result;
  
  const tcStr = String(timeControl);
  
  // Check if correspondence/daily format (1/value)
  if (tcStr.includes('/')) {
    const parts = tcStr.split('/');
    if (parts.length === 2) {
      result.correspondenceTime = parseInt(parts[1]) || null;
    }
  }
  // Check if live format with increment (value+value)
  else if (tcStr.includes('+')) {
    const parts = tcStr.split('+');
    if (parts.length === 2) {
      result.baseTime = parseInt(parts[0]) || null;
      result.increment = parseInt(parts[1]) || null;
    }
  }
  // Simple live format (just value)
  else {
    result.baseTime = parseInt(tcStr) || null;
    result.increment = 0;
  }
  
  return result;
}

// Helper function to process games data
function processGamesData(games, username) {
  const rows = [];
  const derivedRows = [];
  
  // Sort games by timestamp (oldest first) to ensure Last Rating fills correctly
  const sortedGames = games.slice().sort((a, b) => a.end_time - b.end_time);
  
  // Pre-load existing games data once for performance
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  let existingGames = [];
  
  if (gamesSheet && gamesSheet.getLastRow() > 1) {
    const data = gamesSheet.getDataRange().getValues();
    // Build lookup map: format -> array of {timestamp, rating}
    for (let i = 1; i < data.length; i++) {
      try {
        const format = data[i][7]; // Format column (index 7)
        const endDate = data[i][1]; // End Date column
        const endTime = data[i][2]; // End Time column
        const myRating = data[i][8]; // My Rating column (index 8)
        const timestamp = new Date(endDate + ' ' + endTime).getTime() / 1000;
        
        existingGames.push({
          format: format,
          timestamp: timestamp,
          rating: myRating
        });
      } catch (error) {
        // Skip malformed rows
        continue;
      }
    }
  }
  
  for (const game of sortedGames) {
    try {
      if (!game || !game.url || !game.end_time) {
        Logger.log('Skipping game with missing data');
        continue;
      }
      
      const endDate = new Date(game.end_time * 1000);
      const gameId = game.url.split('/').pop();
      const eco = extractECOFromPGN(game.pgn);
      const ecoUrl = extractECOUrlFromPGN(game.pgn);
      const outcome = getGameOutcome(game, username);
      const termination = getGameTermination(game, username);
      const format = getGameFormat(game);
      const timeClass = game.time_class || 'unknown';
      const duration = extractDurationFromPGN(game.pgn);
      
      // Determine my color and opponent
      const isWhite = game.white?.username === CONFIG.USERNAME;
      const myColor = isWhite ? 'white' : 'black';
      const opponent = isWhite ? game.black?.username : game.white?.username;
      const myRating = isWhite ? game.white?.rating : game.black?.rating;
      const oppRating = isWhite ? game.black?.rating : game.white?.rating;
      
      // Calculate Last Rating from pre-loaded data AND games processed in this batch
      let lastRating = null;
      let lastGameTime = 0;
      
      // Check existing games from sheet
      for (const existingGame of existingGames) {
        if (existingGame.format === format && 
            existingGame.timestamp < game.end_time && 
            existingGame.timestamp > lastGameTime) {
          lastGameTime = existingGame.timestamp;
          lastRating = existingGame.rating;
        }
      }
      
      // Parse time control
      const tcParsed = parseTimeControl(game.time_control, game.time_class);
      
      // Extract moves with clocks and times
      const moveData = extractMovesWithClocks(game.pgn, tcParsed.baseTime, tcParsed.increment);
      
      // Create proper date/time objects
      // End Date: Set to midnight of the game's date (no time component)
      const endDateObj = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
      
      // End Time: Create a date with just time component (use epoch date + time)
      const endTimeObj = new Date(1970, 0, 1, endDate.getHours(), endDate.getMinutes(), endDate.getSeconds());
      
      // Combined End DateTime for derived sheet
      const endDateTime = new Date(endDate.getTime());
      
      // Calculate start time from duration
      let startDateTime = null;
      let startDateObj = null;
      let startTimeObj = null;
      
      if (duration && duration > 0) {
        startDateTime = new Date(endDateTime.getTime() - (duration * 1000));
        startDateObj = new Date(startDateTime.getFullYear(), startDateTime.getMonth(), startDateTime.getDate());
        startTimeObj = new Date(1970, 0, 1, startDateTime.getHours(), startDateTime.getMinutes(), startDateTime.getSeconds());
      }
      
      rows.push([
        game.url, endDateObj, endTimeObj, myColor, opponent || 'Unknown',
        outcome, termination, format,
        myRating || 'N/A', oppRating || 'N/A', lastRating || 'N/A',
        gameId, false, false
      ]);
      
      // Calculate Moves (ply count / 2 rounded up)
      const movesCount = moveData.plyCount > 0 ? Math.ceil(moveData.plyCount / 2) : 0;
      
      // Store derived data in hidden sheet
      derivedRows.push([
        gameId,
        game.white?.username || 'Unknown',
        game.black?.username || 'Unknown',
        game.white?.rating || 'N/A',
        game.black?.rating || 'N/A',
        timeClass,
        game.time_control || '',
        tcParsed.type,
        tcParsed.baseTime,
        tcParsed.increment,
        tcParsed.correspondenceTime,
        eco,
        ecoUrl,
        game.rated !== undefined ? game.rated : true,
        endDateTime,
        startDateTime,
        startDateObj,
        startTimeObj,
        duration,
        moveData.plyCount,
        movesCount,
        moveData.moveList,
        moveData.clocks,
        moveData.times
      ]);
      
      // Add this game to existingGames for subsequent games in this batch
      existingGames.push({
        format: format,
        timestamp: game.end_time,
        rating: myRating
      });
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  // Write derived data to hidden sheet
  if (derivedSheet && derivedRows.length > 0) {
    const lastRow = derivedSheet.getLastRow();
    derivedSheet.getRange(lastRow + 1, 1, derivedRows.length, derivedRows[0].length).setValues(derivedRows);
  }
  
  return rows;
}

// Get game format based on rules and time control
function getGameFormat(game) {
  const rules = game.rules || 'chess';
  const timeClass = game.time_class || 'unknown';
  
  if (rules === 'chess') {
    return timeClass;  // bullet, blitz, rapid, daily
  } else if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  } else {
    return rules;
  }
}

// Remove duplicate games based on Game ID
function removeDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Games sheet not found!');
    return;
  }
  
  const data = gamesSheet.getDataRange().getValues();
  const header = data[0];
  const gameIdCol = 11; // Game ID column (index 11)
  
  const seen = new Set();
  const rowsToKeep = [header];
  let duplicateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const gameId = data[i][gameIdCol];
    
    if (!seen.has(gameId)) {
      seen.add(gameId);
      rowsToKeep.push(data[i]);
    } else {
      duplicateCount++;
    }
  }
  
  if (duplicateCount > 0) {
    gamesSheet.clear();
    gamesSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    
    gamesSheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Removed ${duplicateCount} duplicate(s)`, 
      '🗑️', 
      3
    );
    Logger.log(`Removed ${duplicateCount} duplicates`);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No duplicates found!', 'ℹ️', 2);
  }
}
// ============================================
// HELPER FUNCTIONS
// ============================================

function extractMovesFromPGN(pgn) {
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  
  return moveSection
    .replace(/\{[^}]*\}/g, '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\$\d+/g, '')
    .replace(/\d+\.{3}/g, '')
    .replace(/\d+\./g, '')
    .replace(/[!?+#]+/g, '')
    .trim()
    .split(/\s+/)
    .filter(m => m && m !== '*' && !m.match(/^(1-0|0-1|1\/2-1\/2)$/));
}

function extractECOFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECO "([^"]+)"\]/);
  return match ? match[1] : '';
}

// Extract ECO URL from PGN
function extractECOUrlFromPGN(pgn) {
  if (!pgn) return '';
  const match = pgn.match(/\[ECOUrl "([^"]+)"\]/);
  return match ? match[1] : '';
}

// Extract moves with clock times from PGN
function extractMovesWithClocks(pgn, baseTime, increment) {
  if (!pgn) return { moves: [], clocks: [], times: [] };
  
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  const moves = [];
  const clocks = [];
  const times = [];
  
  // Regex to match move and its clock: "e4 {[%clk 0:02:59.9]}"
  const movePattern = /([NBRQK]?[a-h]?[1-8]?x?[a-h][1-8](?:=[NBRQK])?|O-O(?:-O)?)\s*\{?\[%clk\s+(\d+):(\d+):(\d+)(?:\.(\d+))?\]?\}?/g;
  
  let match;
  let prevClock = [baseTime || 0, baseTime || 0]; // [white, black] previous clocks
  let moveIndex = 0;
  
  while ((match = movePattern.exec(moveSection)) !== null) {
    const move = match[1];
    const hours = parseInt(match[2]) || 0;
    const minutes = parseInt(match[3]) || 0;
    const seconds = parseInt(match[4]) || 0;
    const deciseconds = parseInt(match[5]) || 0;
    
    // Convert clock to total seconds
    const clockSeconds = hours * 3600 + minutes * 60 + seconds + deciseconds / 10;
    
    moves.push(move);
    clocks.push(clockSeconds);
    
    // Calculate time spent on this move
    const playerIndex = moveIndex % 2; // 0 = white, 1 = black
    const prevPlayerClock = prevClock[playerIndex];
    
    // Time spent = previous clock - current clock + increment
    let timeSpent = prevPlayerClock - clockSeconds + (increment || 0);
    
    // Minimum move time is 0.1 seconds (Chess.com enforces this)
    if (timeSpent < 0.1) timeSpent = 0.1;
    
    times.push(Math.round(timeSpent * 10) / 10); // Round to 1 decimal
    
    // Update previous clock for this player
    prevClock[playerIndex] = clockSeconds;
    
    moveIndex++;
  }
  
  return { 
    moveList: moves.join(', '), 
    clocks: clocks.join(', '), 
    times: times.join(', '),
    plyCount: moves.length
  };
}

function extractDurationFromPGN(pgn) {
  if (!pgn) return null;
  
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  const endDateMatch = pgn.match(/\[EndDate "([^"]+)"\]/);
  const endTimeMatch = pgn.match(/\[EndTime "([^"]+)"\]/);
  
  if (!dateMatch || !timeMatch || !endDateMatch || !endTimeMatch) {
    return null;
  }
  
  try {
    const startDateParts = dateMatch[1].split('.');
    const startTimeParts = timeMatch[1].split(':');
    const startDate = new Date(Date.UTC(
      parseInt(startDateParts[0]),
      parseInt(startDateParts[1]) - 1,
      parseInt(startDateParts[2]),
      parseInt(startTimeParts[0]),
      parseInt(startTimeParts[1]),
      parseInt(startTimeParts[2])
    ));
    
    const endDateParts = endDateMatch[1].split('.');
    const endTimeParts = endTimeMatch[1].split(':');
    const endDate = new Date(Date.UTC(
      parseInt(endDateParts[0]),
      parseInt(endDateParts[1]) - 1,
      parseInt(endDateParts[2]),
      parseInt(endTimeParts[0]),
      parseInt(endTimeParts[1]),
      parseInt(endTimeParts[2])
    ));
    
    const durationMs = endDate.getTime() - startDate.getTime();
    return Math.round(durationMs / 1000);
  } catch (error) {
    Logger.log(`Error parsing duration: ${error.message}`);
    return null;
  }
}

function getGameOutcome(game, username) {
  if (!game || !game.white || !game.black) return 'unknown';
  
  const isWhite = game.white?.username === CONFIG.USERNAME;
  const myResult = isWhite ? game.white.result : game.black.result;
  
  if (!myResult) return 'unknown';
  
  return myResult || 'unknown';
}

function getGameTermination(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white?.username === CONFIG.USERNAME;
  const myResult = isWhite ? game.white.result : game.black.result;
  const opponentResult = isWhite ? game.black.result : game.white.result;
  
  if (!myResult) return 'Unknown';
  
  // If I won, use opponent's result for termination
  if (myResult === 'win') {
    return opponentResult || 'unknown';
  }
  
}

// ============================================
// MAIN MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('♟️ Chess Analyzer')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addItem('2️⃣ Initial Fetch (All Games)', 'fetchAllGamesInitialOptimized')
    .addItem('3️⃣ Update Recent Games', 'fetchChesscomGamesOptimized')
    .addSeparator()
    .addItem('📋 Fetch Callback Last 10', 'fetchCallbackLast10')
    .addSeparator()
    .addItem('📊 Create/Update Daily Stats', 'setupDailyStatsSheet')
    .addToUi();
}
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  if (!gamesSheet) {
    gamesSheet = ss.insertSheet(SHEETS.GAMES);
    const headers = [
      'Game URL', 'End Date', 'End Time', 'My Color', 'Opponent',
      'Outcome', 'Termination', 'Format',
      'My Rating', 'Opp Rating', 'Last Rating',
      'Game ID', 'Analyzed', 'Callback Fetched'
    ];
    gamesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    gamesSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    gamesSheet.setFrozenRows(1);
    gamesSheet.setColumnWidth(1, 200);
    
    // Format date and time columns
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
    
    // Add conditional formatting for My Color
    const colorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('white')
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setRanges([gamesSheet.getRange('D2:D')])
      .build();
    const colorRule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('black')
      .setBackground('#333333')
      .setFontColor('#FFFFFF')
      .setRanges([gamesSheet.getRange('D2:D')])
      .build();
    
    // Conditional formatting for Outcome
    const winRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('win')
      .setBackground('#d9ead3')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    const lossRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('loss')
      .setBackground('#f4cccc')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    const drawRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('draw')
      .setBackground('#fff2cc')
      .setRanges([gamesSheet.getRange('F2:F')])
      .build();
    
    // Conditional formatting for booleans
    const trueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=M2=TRUE')
      .setBackground('#d9ead3')
      .setRanges([gamesSheet.getRange('M2:N')])
      .build();
    
    const rules = gamesSheet.getConditionalFormatRules();
    rules.push(colorRule, colorRule2, winRule, lossRule, drawRule, trueRule);
    gamesSheet.setConditionalFormatRules(rules);
  }
  
  let derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  if (!derivedSheet) {
    derivedSheet = ss.insertSheet(SHEETS.DERIVED);
    const headers = [
      'Game ID', 'White Username', 'Black Username', 'White Rating', 'Black Rating',
      'Time Class', 'Time Control', 'Type', 'Base Time', 'Increment', 'Correspondence Time',
      'ECO', 'ECO URL', 'Rated',
      'End', 'Start', 'Start Date', 'Start Time', 'Duration (s)', 'Ply Count', 'Moves',
      'Move List', 'Move Clocks', 'Move Times'
    ];
    derivedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    derivedSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#666666')
      .setFontColor('#ffffff');
    derivedSheet.setFrozenRows(1);
    
    // Format date/time columns
    derivedSheet.getRange('O:O').setNumberFormat('m"/"d"/"yy h:mm AM/PM'); // End
    derivedSheet.getRange('P:P').setNumberFormat('m"/"d"/"yy h:mm AM/PM'); // Start
    derivedSheet.getRange('Q:Q').setNumberFormat('m"/"d"/"yy'); // Start Date
    derivedSheet.getRange('R:R').setNumberFormat('h:mm AM/PM'); // Start Time
    
    // Format Move Clocks and Move Times columns as text to prevent scientific notation
    derivedSheet.getRange('V:V').setNumberFormat('@STRING@'); // Move List
    derivedSheet.getRange('W:W').setNumberFormat('@STRING@'); // Move Clocks
    derivedSheet.getRange('X:X').setNumberFormat('@STRING@'); // Move Times
    
    // Hide the derived sheet
    derivedSheet.hideSheet();
  }
  
  // Apply formatting to existing Games sheet if it exists
  if (gamesSheet) {
    gamesSheet.getRange('B:B').setNumberFormat('m"/"d"/"yy');
    gamesSheet.getRange('C:C').setNumberFormat('h:mm AM/PM');
  }
  
  
  let callbackSheet = ss.getSheetByName(SHEETS.CALLBACK);
  if (!callbackSheet) {
    callbackSheet = ss.insertSheet(SHEETS.CALLBACK);
    const headers = [
      'Game ID', 'Game URL', 'Callback URL', 'End Time', 'My Color', 'Time Class',
      'My Rating', 'Opp Rating', 'My Rating Change', 'Opp Rating Change',
      'My Rating Before', 'Opp Rating Before',
      'Base Time', 'Time Increment', 'Move Timestamps',
      'My Username', 'My Country', 'My Membership', 'My Member Since',
      'My Default Tab', 'My Post Move Action', 'My Location',
      'Opp Username', 'Opp Country', 'Opp Membership', 'Opp Member Since',
      'Opp Default Tab', 'Opp Post Move Action', 'Opp Location',
      'Date Fetched'
    ];
    callbackSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    callbackSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f4b400')
      .setFontColor('#ffffff');
    callbackSheet.setFrozenRows(1);
    
    // Format Move Timestamps column as text to prevent scientific notation
    callbackSheet.getRange('K:K').setNumberFormat('@STRING@');
  }
  
  // Setup Daily Stats sheet
  let dailySheet = ss.getSheetByName('Daily Stats');
  if (!dailySheet) {
    dailySheet = ss.insertSheet('Daily Stats');
    const headers = ['Date'];
    dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dailySheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    dailySheet.setFrozenRows(1);
    dailySheet.setFrozenColumns(1);
    
    // Format date column
    dailySheet.getRange('A:A').setNumberFormat('m"/"d"/"yy');
  }
  
  SpreadsheetApp.getUi().alert('✅ Sheets setup complete!');
}
  
  

// ============================================
//  DAILY STATS SHEET
// ============================================

function setupDailyStatsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const derivedSheet = ss.getSheetByName(SHEETS.DERIVED);
  
  if (!gamesSheet || !derivedSheet) {
    SpreadsheetApp.getUi().alert('❌ Please run "Setup Sheets" and fetch games first!');
    return;
  }
  
  let dailySheet = ss.getSheetByName('Daily Stats');
  const isNewSheet = !dailySheet;
  
  if (isNewSheet) {
    dailySheet = ss.insertSheet('Daily Stats');
  }
  
  const gamesData = gamesSheet.getDataRange().getValues();
  if (gamesData.length <= 1) {
    SpreadsheetApp.getUi().alert('❌ No games found! Please fetch games first.');
    return;
  }
  
  // Get unique formats
  const formats = [...new Set(gamesData.slice(1).map(row => row[7]))].sort();
  
  // Build headers
  const headers = ['Date'];
  for (const format of formats) {
    headers.push(`${format} wins`, `${format} losses`, `${format} draws`, `${format} Duration (s)`, `${format} Rating`);
  }
  
  // Get last processed date from sheet (or null if new)
  let lastProcessedDate = null;
  if (!isNewSheet && dailySheet.getLastRow() > 1) {
    const lastRow = dailySheet.getLastRow();
    lastProcessedDate = new Date(dailySheet.getRange(lastRow, 1).getValue());
  }
  
  // Get min and max dates from games
  const dates = gamesData.slice(1).map(row => new Date(row[1])).filter(d => !isNaN(d));
  const minDate = new Date(Math.min(...dates));
  const maxDate = new Date(Math.max(...dates));
  
  // Determine which dates to process
  const startDate = lastProcessedDate ? new Date(lastProcessedDate.getTime() + 86400000) : minDate; // +1 day
  
  if (startDate > maxDate) {
    ss.toast('No new dates to process!', 'ℹ️', 2);
    return;
  }
  
  // Generate dates to process
  const datesToProcess = [];
  const currentDate = new Date(startDate);
  while (currentDate <= maxDate) {
    datesToProcess.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  // Build duration lookup
  const derivedData = derivedSheet.getDataRange().getValues();
  const durationByGameId = new Map();
  for (let i = 1; i < derivedData.length; i++) {
    const gameId = derivedData[i][0];
    const duration = derivedData[i][18];
    if (gameId && duration) durationByGameId.set(gameId, duration);
  }
  
  // Build games by date map (only for dates we're processing)
  const gamesByDate = new Map();
  for (let i = 1; i < gamesData.length; i++) {
    const endDate = new Date(gamesData[i][1]);
    if (endDate < startDate || endDate > maxDate) continue; // Skip dates outside range
    
    const dateKey = endDate.toDateString();
    const format = gamesData[i][7];
    const outcome = gamesData[i][5];
    const rating = gamesData[i][8];
    const gameId = gamesData[i][11];
    const duration = durationByGameId.get(gameId) || 0;
    
    if (!gamesByDate.has(dateKey)) gamesByDate.set(dateKey, []);
    gamesByDate.get(dateKey).push({
      format: format,
      outcome: outcome,
      rating: rating,
      duration: duration,
      timestamp: endDate.getTime()
    });
  }
  
  // Get last known ratings (from last row of existing sheet, or initialize)
  const lastRating = {};
  if (!isNewSheet && dailySheet.getLastRow() > 1) {
    const lastRow = dailySheet.getLastRow();
    const lastRowData = dailySheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
    
    for (let i = 0; i < formats.length; i++) {
      const colOffset = 1 + (i * 5) + 4; // Rating column for each format
      lastRating[formats[i]] = lastRowData[colOffset];
    }
  } else {
    formats.forEach(f => lastRating[f] = null);
  }
  
  // Build new rows
  const newRows = [];
  
  // Add header if new sheet
  if (isNewSheet) {
    newRows.push(headers);
  }
  
  // Process each date
  for (const date of datesToProcess) {
    const dateKey = date.toDateString();
    const row = [date];
    const gamesOnDate = gamesByDate.get(dateKey) || [];
    
    gamesOnDate.sort((a, b) => a.timestamp - b.timestamp);
    
    for (const format of formats) {
      const formatGames = gamesOnDate.filter(g => g.format === format);
      
      const wins = formatGames.filter(g => g.outcome === 'win').length;
      const losses = formatGames.filter(g => g.outcome === 'loss').length;
      const draws = formatGames.filter(g => g.outcome === 'draw').length;
      const totalDuration = formatGames.reduce((sum, g) => sum + g.duration, 0);
      
      // Update rating if games played
      if (formatGames.length > 0) {
        const lastGame = formatGames[formatGames.length - 1];
        lastRating[format] = lastGame.rating !== 'N/A' ? lastGame.rating : lastRating[format];
      }
      
      row.push(wins || null, losses || null, draws || null, totalDuration || null, lastRating[format]);
    }
    
    newRows.push(row);
  }
  
  // Write to sheet
  const startRow = isNewSheet ? 1 : dailySheet.getLastRow() + 1;
  dailySheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  
  
  ss.toast(`✅ Added ${newRows.length - (isNewSheet ? 1 : 0)} days to Daily Stats!`, '✅', 3);
}
