/**
 * @fileoverview GameParser.gs
 * @author YourName
 * @license MIT
 * @description Provides utility functions for parsing and transforming
 *              individual Game objects retrieved from the Chess.com API,
 *              with a focus on "my vs. opponent" perspective.
 * @see https://www.chess.com/announcements/view/published-data-api
 */

var GameParser = (function() {

  /**
   * Transforms a raw Game object into a more user-friendly or specific format,
   * establishing "my" and "opponent" variables based on the provided `perspectiveUsername`.
   *
   * @param {object} rawGameObject The raw Game object as returned by the Chess.com API.
   * @param {string} perspectiveUsername The username of the player whose perspective we are parsing from ("me").
   * @return {object|null} A processed game object with selected or formatted properties,
   *                       including "my" and "opp" prefixed variables. Returns null if invalid.
   */
  function parseGame(rawGameObject, perspectiveUsername) {
    if (!rawGameObject || !perspectiveUsername) {
      Logger.log("GameParser.parseGame received null/undefined game object or perspectiveUsername.");
      return null;
    }

    // --- Common Game Variables ---
    var gameUrl = rawGameObject.url || null;
    var pgnContent = rawGameObject.pgn || null;
    var fen = rawGameObject.fen || null;
    var timeControl = rawGameObject.time_control || null;
    var timeClass = rawGameObject.time_class || null;
    var rules = rawGameObject.rules || "chess";
    var rated = rawGameObject.rated !== undefined ? rawGameObject.rated : null;
    var eco = rawGameObject.eco || null;
    var tournament = rawGameObject.tournament || null;
    var match = rawGameObject.match || null;
    
    // Parse dates
    var startDate = rawGameObject.start_time ? new Date(rawGameObject.start_time * 1000) : null;
    var endDate = rawGameObject.end_time ? new Date(rawGameObject.end_time * 1000) : null;

    var whitePlayer = rawGameObject.white;
    var blackPlayer = rawGameObject.black;
    
    // Validate player data exists
    if (!whitePlayer || !blackPlayer) {
      Logger.log("Missing player data in game object.");
      return null;
    }

    // Derive common result format (e.g., "1-0", "0-1", "1/2-1/2")
    var result = null;
    if (whitePlayer.result === "win") {
      result = "1-0";
    } else if (blackPlayer.result === "win") {
      result = "0-1";
    } else if (isDrawResult(whitePlayer.result) || isDrawResult(blackPlayer.result)) {
      result = "1/2-1/2";
    }

    // --- "Me vs. Opponent" Variables ---
    var isWhite = whitePlayer.username.toLowerCase() === perspectiveUsername.toLowerCase();
    var isBlack = blackPlayer.username.toLowerCase() === perspectiveUsername.toLowerCase();
    
    if (!isWhite && !isBlack) {
      Logger.log("Perspective username '" + perspectiveUsername + "' not found in game.");
      return null;
    }

    var myPlayer = isWhite ? whitePlayer : blackPlayer;
    var oppPlayer = isWhite ? blackPlayer : whitePlayer;
    var myColor = isWhite ? "white" : "black";
    var oppColor = isWhite ? "black" : "white";
    
    var accuracies = rawGameObject.accuracies;
    
    var myUsername = myPlayer.username;
    var myRating = myPlayer.rating || null;
    var myResult = myPlayer.result || null;
    var myAccuracy = accuracies ? accuracies[myColor] : null;
    var myUuid = myPlayer.uuid || null;
    var myProfileUrl = myPlayer["@id"] || null;
    
    var oppUsername = oppPlayer.username;
    var oppRating = oppPlayer.rating || null;
    var oppResult = oppPlayer.result || null;
    var oppAccuracy = accuracies ? accuracies[oppColor] : null;
    var oppUuid = oppPlayer.uuid || null;
    var oppProfileUrl = oppPlayer["@id"] || null;

    return {
      // Game metadata
      gameUrl: gameUrl,
      pgnContent: pgnContent,
      fen: fen,
      timeControl: timeControl,
      timeClass: timeClass,
      rules: rules,
      rated: rated,
      eco: eco,
      tournament: tournament,
      match: match,
      startDate: startDate,
      endDate: endDate,
      result: result,

      // My perspective
      myUsername: myUsername,
      myRating: myRating,
      myResult: myResult,
      myAccuracy: myAccuracy,
      myUuid: myUuid,
      myProfileUrl: myProfileUrl,
      playedAs: myColor,

      // Opponent's perspective
      oppUsername: oppUsername,
      oppRating: oppRating,
      oppResult: oppResult,
      oppAccuracy: oppAccuracy,
      oppUuid: oppUuid,
      oppProfileUrl: oppProfileUrl
    };
  }

  /**
   * Helper function to determine if a result code represents a draw.
   * @param {string} resultCode - The result code from the Chess.com API.
   * @return {boolean} True if the result is a draw.
   */
  function isDrawResult(resultCode) {
    var drawResults = [
      "agreed",
      "repetition", 
      "stalemate",
      "insufficient",
      "50move",
      "timevsinsufficient"
    ];
    return drawResults.indexOf(resultCode) !== -1;
  }

  /**
   * Processes an array of raw Game objects using the parseGame function,
   * for a specific user's perspective.
   *
   * @param {object[]|null} rawGameObjectsArray An array of raw Game objects.
   * @param {string} perspectiveUsername The username of the player whose perspective ("me") is used for parsing.
   * @return {object[]|null} An array of processed game objects, or null if input is invalid.
   */
  function parseGamesArray(rawGameObjectsArray, perspectiveUsername) {
    if (!rawGameObjectsArray || !Array.isArray(rawGameObjectsArray)) {
      Logger.log("GameParser.parseGamesArray received invalid input (not an array).");
      return null;
    }
    if (!perspectiveUsername) {
      Logger.log("GameParser.parseGamesArray received null/undefined perspectiveUsername.");
      return null;
    }

    return rawGameObjectsArray
      .map(function(game) {
        return parseGame(game, perspectiveUsername);
      })
      .filter(Boolean);
  }

  /**
   * Parses the Chess.com API response which wraps games in a "games" array.
   * @param {object} apiResponse - The full API response with { games: [...] }
   * @param {string} perspectiveUsername - The username of the player whose perspective to use.
   * @return {object[]|null} An array of processed game objects, or null if invalid.
   */
  function parseApiResponse(apiResponse, perspectiveUsername) {
    if (!apiResponse || !apiResponse.games) {
      Logger.log("GameParser.parseApiResponse received invalid API response (missing 'games' property).");
      return null;
    }
    return parseGamesArray(apiResponse.games, perspectiveUsername);
  }

  /**
   * Extract ECO code from ECO URL if present.
   * @param {string} ecoUrl - The ECO URL from the API.
   * @return {string|null} The ECO code (e.g., "B00") or null.
   */
  function extractEcoCode(ecoUrl) {
    if (!ecoUrl) return null;
    var parts = ecoUrl.split('/');
    return parts[parts.length - 1] || null;
  }

  /**
   * Get a human-readable description of a result code.
   * @param {string} resultCode - The result code from Chess.com API.
   * @return {string} Human-readable description.
   */
  function getResultDescription(resultCode) {
    var descriptions = {
      "win": "Win",
      "checkmated": "Checkmated",
      "agreed": "Draw agreed",
      "repetition": "Draw by repetition",
      "timeout": "Timeout",
      "resigned": "Resigned",
      "stalemate": "Stalemate",
      "lose": "Lose",
      "insufficient": "Insufficient material",
      "50move": "Draw by 50-move rule",
      "abandoned": "Abandoned",
      "kingofthehill": "King reached the hill",
      "threecheck": "Checked 3 times",
      "timevsinsufficient": "Timeout vs insufficient material",
      "bughousepartnerlose": "Bughouse partner lost"
    };
    return descriptions[resultCode] || resultCode;
  }

  /**
   * Determines if the user won the game.
   * @param {string} myResult - The user's result code.
   * @return {boolean|null} True if won, false if lost, null if draw.
   */
  function didIWin(myResult) {
    if (myResult === "win") return true;
    if (myResult === "lose" || myResult === "checkmated" || 
        myResult === "timeout" || myResult === "resigned" || 
        myResult === "abandoned") return false;
    return null; // Draw or other
  }

  return {
    parseGame: parseGame,
    parseGamesArray: parseGamesArray,
    parseApiResponse: parseApiResponse,
    isDrawResult: isDrawResult,
    extractEcoCode: extractEcoCode,
    getResultDescription: getResultDescription,
    didIWin: didIWin
  };
})();
