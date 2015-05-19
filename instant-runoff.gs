/*
First Passed The Post and Instant Runoff Voting with Google Form and Google Apps Script
Author: Darrell Ross

This project may contain bugs. Use at your own risk.

This script has been designed for interaction straight from the Spreadsheet.
There should be no need to edit this file at all.

Instructions: https://docs.google.com/a/milsoft.com/document/d/1bLEOPwxfSTwEh_pQglM7nGkUtaFkjGu5QWqEV17FhcE/edit?usp=sharing

TODO
 - eliminate dependence on the global variable VOTE_TYPE_NAME

*/

/* BEGIN SETTINGS */ 
/**************************/
var VOTE_SHEET_NAME = "Votes";                   // active sheet where form data is entered will be renamed to this
var CONFIGURE_SHEET_NAME = "Configure";          // name of the configure sheet name where you enter valid keys to use and designate votes and choice counts
var RESULTS_SHEET_NAME = "Results";              // default value. when run without special function, Used Keys has no prefix
var BASE_ROW = 2;                                // row where votes start - could change it if you wanted to keep old votes
/**************************/
var VOTE_TYPE_NAME;                              // global variable indicating which vote is being run
/**************************/
/* END SETTINGS */ 

function VoteType(pvote,pindex,pchoices,pbasecol)
{
  this.VoteName = pvote;
  this.VoteIndex = pindex;
  this.ChoiceCount = pchoices;
  this.BaseColumn = pbasecol;
}
/**************************/
function WinnerData(pname,pvote,pdate)
{
  this.WinnerName = pname;
  this.WinnerVote = pvote;
  this.WinnerDate = pdate;
}
/**************************/
var USING_KEYS = true;        // must be initted using InitUsingKeys()
var KEYS_COLUMN = 2;          // column where keys would appear
function InitUsingKeys() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if(active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME) == null)
    SetupConfigureSheet();  
  var usingKeys = active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("A1").getNote();
  if(usingKeys == "yes")
    USING_KEYS = true;  
  else
    USING_KEYS = false;
}
/**************************/
var VOTE_TYPE_COUNT = 0;                     // Number of vote types - set during InitVoteTypesArray();
var VOTE_TYPE_ARRAY = [];
var WINNER_ARRAY = [];                       // All winner data
/**************************/
// Look up all votes //
function InitVoteTypesArray() {
  var firstRow = 2;          // skip header
  var voteColumn = 3;        // vote name is in column C
  var choiceCountColumn = 4; // choice count is in column D
  var numColumns = 2;        // vote and choicecount columns
  var results_range = get_range_with_values(CONFIGURE_SHEET_NAME, firstRow, voteColumn, numColumns);
  
  if (results_range == null) {
    Browser.msgBox("No vote positions listed. Looking within sheet: " + CONFIGURE_SHEET_NAME);
    return false;
  }

  var baseCol = 3;
  if(!USING_KEYS)
    baseCol = 2;
  var previousChoiceCount = 0;
  VOTE_TYPE_COUNT = results_range.getNumRows();
  for (var row = 1; row <= VOTE_TYPE_COUNT; row++) {
    var voteTypeCell = results_range.getCell(row, 1);    // Vote Type is First Column of results_range
    var voteTypeCellValue = voteTypeCell.getValue();
    var choiceCountCell = results_range.getCell(row, 2); // Vote Choice Count is Second Column of results_range
    var choiceCountCellValue = choiceCountCell.getValue();
    baseCol += previousChoiceCount;
    previousChoiceCount = choiceCountCellValue;
    VOTE_TYPE_ARRAY.push(new VoteType(voteTypeCellValue, row, choiceCountCellValue, baseCol));
  }  
  return true;
}
/**************************/
function GetVoteIndex(voteTypeName) {
  if(typeof(voteTypeName)==='undefined')
    return 0;
  
  for (var i = 0; i < VOTE_TYPE_COUNT; i++) {
    if(VOTE_TYPE_ARRAY[i].VoteName == voteTypeName)
      return VOTE_TYPE_ARRAY[i].VoteIndex;
  }
  return 0;
}
/**************************/
function GetChoiceCount(voteTypeName) {
  if(typeof(voteTypeName)==='undefined')
    return 0;
  
  for (var i = 0; i < VOTE_TYPE_COUNT; i++) {
    if(VOTE_TYPE_ARRAY[i].VoteName == voteTypeName)
      return VOTE_TYPE_ARRAY[i].ChoiceCount;
  }
  return 0;
}
/**************************/
function GetBaseColumn(voteTypeName) {
  if(typeof(voteTypeName)==='undefined')
    return 0;
  
  for (var i = 0; i < VOTE_TYPE_COUNT; i++) {
    if(VOTE_TYPE_ARRAY[i].VoteName == voteTypeName)
      return VOTE_TYPE_ARRAY[i].BaseColumn;
  }
  return 0;
}
/**************************/
/**************************/
/**************************/


function initialize_spreadsheet() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  active_spreadsheet.getSheets()[0].setName(VOTE_SHEET_NAME); // Rename Form Entries Sheet to "Votes"
  SetupConfigureSheet(); // Initialize Configuration Sheet
  create_menu_items();
}

// Clears away old tallied votes. Does not remove configuration.
function clear_voting() {
  initialize_spreadsheet(); // just to be safe
  InitVoteTypesArray();
  SetupResultsSheet();
  ResetBackgroundColors();  
}

// Tallies the votes
function tally_votes() {
  clear_voting(); // start clean
  InitUsingKeys();
  InitVoteTypesArray();
  var success = true;
  for(var i=0; i<VOTE_TYPE_COUNT; i++)
  {
    VOTE_TYPE_NAME = VOTE_TYPE_ARRAY[i].VoteName; // global variable used to access all VOTE_TYPE_ARRAY data
    if(!tally_single_vote(VOTE_TYPE_ARRAY[i].VoteName)) {
      success = false;
      break;
    }
  }
  if(success)
    OutputVoteResults();
  else
    Browser.msgBox("Vote Tallying failed.");
}


/* Notification state */
var missing_keys_used_sheet_alert = false;
/* End notification state */

// Creates the Configure Sheet if it doesn't exist. Does not remove configuration data.
function SetupConfigureSheet() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME) == null) {
    active_spreadsheet.insertSheet(CONFIGURE_SHEET_NAME);
    active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("A1").setNote("yes");
  }

  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("A1").setValue("Keys").setFontWeight("bold");
  var keysHelpText = "Enter the keys starting on the second row. One key per cell.";
  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("A2").setNote(keysHelpText);
  
  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("C1").setValue("Votes").setFontWeight("bold");
  var votesHelpText = "Enter the names of each vote you are holding. Enter them in the same order as you have them on your form.";
  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("C1").setNote(votesHelpText).setFontWeight("bold");

  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("D1").setValue("Choice Counts").setFontWeight("bold");
  var choicesHelpText = "Enter the quantity of choices you have. If this is First Passed The Post voting, then all entries will be 1. If it is Instant-Runoff, enter the number of choices for each item.";
  active_spreadsheet.getSheetByName(CONFIGURE_SHEET_NAME).getRange("D1").setNote(choicesHelpText).setFontWeight("bold");
}

// If the RESULTS_SHEET_NAME Sheet Exists, then we delete it first. If it does not exist, we create it.
function SetupResultsSheet() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (active_spreadsheet.getSheetByName(RESULTS_SHEET_NAME) == null)
    active_spreadsheet.insertSheet(RESULTS_SHEET_NAME);

  // Clear and Repopulate Column Headers //
  active_spreadsheet.getSheetByName(RESULTS_SHEET_NAME).clear();
  active_spreadsheet.getSheetByName(RESULTS_SHEET_NAME).clearNotes();
  for (var i=0; i<VOTE_TYPE_COUNT; i++) {
    var range = String.fromCharCode(65+i) + '1';
    active_spreadsheet.getSheetByName(RESULTS_SHEET_NAME).getRange(range).setValue(VOTE_TYPE_ARRAY[i].VoteName).setFontWeight("bold");
  }
}

function OutputVoteResults() {
  var fullResultString = "";
  for (var i=0; i<VOTE_TYPE_COUNT; i++) {
    var winner = WINNER_ARRAY[i].WinnerName;
    var winnerVote = WINNER_ARRAY[i].WinnerVote;
    var winnerDate = WINNER_ARRAY[i].WinnerDate;
    var winnerVoteIndex = GetVoteIndex(WINNER_ARRAY[i].WinnerVote)-1;
    var winnerResultRangeString = String.fromCharCode(65+winnerVoteIndex) + '1';
    var winnerResultCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME).getRange(winnerResultRangeString).getCell(1,1);
    var winnerMessage = "Winner: " + winner + winnerDate;
    winnerResultCell.setNote(winnerMessage);    
    fullResultString += "|****|" + winnerVote + " " + winnerMessage;
  }
  Browser.msgBox(fullResultString);
}

/***** MENU CONFIGURATION *****/
function create_menu_items() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [ {name: "Initialize Spreadsheet", functionName: "initialize_spreadsheet"},
                        {name: "Setup Voting", functionName: "clear_voting"},
                        {name: "Tally Votes", functionName: "tally_votes"}
                      ];
    ss.addMenu("VOTING", menuEntries);
}
/* Create menus */
function onOpen() {
    initialize_spreadsheet();
}
/* Create menus when installed */
function onInstall() {
    onOpen();
}
/***** END MENU CONFIGURATION *****/

function tally_single_vote(voteTypeName) {                       
  /* Determine number of voting columns */
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var row1_range = active_spreadsheet.getSheetByName(VOTE_SHEET_NAME).getRange("A1:1");
  var choiceColumnCount = GetChoiceCount(voteTypeName);
  var baseColumn = GetBaseColumn(voteTypeName);

  /* Reset state */
  missing_keys_used_sheet_alert = false;
  
  /* Begin */
  clear_background_color(baseColumn, choiceColumnCount);
  
  var results_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, baseColumn, choiceColumnCount);
  
  if (results_range == null) {
    Browser.msgBox("No votes. Looking for sheet: " + VOTE_SHEET_NAME);
    return false;
  }
  // Keys are used to prevent voters from voting twice.
  // Keys also allow voters to change their vote.
  // If keys_range == null then we are not using keys. 
  var keys_range = null;
  
  // List of valid keys
  var valid_keys;
  
  if (USING_KEYS) {
    keys_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, KEYS_COLUMN, 1);
    if (keys_range == null) {
      Browser.msgBox("Using keys and could not find column with submitted keys. " + 
                     "Looking in column " + KEYS_COLUMN + 
                     " in sheet: " + VOTE_SHEET_NAME);
      return false;
    }
    var valid_keys_range = get_range_with_values(CONFIGURE_SHEET_NAME, BASE_ROW, 1, 1);
    if (valid_keys_range == null) {
      var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURE_SHEET_NAME);
      if (results_sheet == null) {
        Browser.msgBox("Looking for list of valid keys. Cannot find sheet: " + CONFIGURE_SHEET_NAME);
      } else {
        Browser.msgBox("List of valid keys cannot be found in sheet: " + CONFIGURE_SHEET_NAME);
      }
      return false;
    }
    valid_keys = range_to_array(valid_keys_range);
  }
  
  /* candidates is a list of names (strings) */
  var candidates = get_all_candidates(results_range);
  
  /* votes is an object mapping candidate names -> number of votes */
  var votes = get_votes(results_range, candidates, keys_range, valid_keys);
  
  /* winner is candidate name (string) or null */
  var winner = get_winner(votes, candidates);


  while (winner == null) {
    /* Modify candidates to only include remaining candidates */
    get_remaining_candidates(votes, candidates);
    if (candidates.length == 0) {
      if (missing_keys_used_sheet_alert) {
        Browser.msgBox("Unable to record keys used. Looking for sheet: " + RESULTS_SHEET_NAME);    
      }
      var dateTimeMessage = " \nDate and time: " + Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
      WINNER_ARRAY.push(new WinnerData("TIE",VOTE_TYPE_NAME,dateTimeMessage));
      return true;
    }
    votes = get_votes(results_range, candidates, keys_range, valid_keys);
    winner = get_winner(votes, candidates);
  }
  
  if (missing_keys_used_sheet_alert) {
    Browser.msgBox("Unable to record keys used. Looking for sheet: " + RESULTS_SHEET_NAME);    
  }
  var dateTimeMessage = " \nDate and time: " + Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
  WINNER_ARRAY.push(new WinnerData(winner,VOTE_TYPE_NAME,dateTimeMessage));
  return true;
}


function get_range_with_values(sheet_string, base_row, base_column, num_columns) {
  var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_string);
  if (results_sheet == null) {
    return null;
  }
  var a1string = String.fromCharCode(65 + base_column - 1) +
      base_row + ':' + 
      String.fromCharCode(65 + base_column + num_columns - 2);
  var results_range = results_sheet.getRange(a1string);
  // results_range contains the whole columns all the way to
  // the bottom of the spreadsheet. We only want the rows
  // with votes in them, so we're going to count how many
  // there are and then just return those.
  var num_rows = get_num_rows_with_values(results_range);
  if (num_rows == 0) {
    return null;
  }
  results_range = results_sheet.getRange(base_row, base_column, num_rows, num_columns);
  return results_range;
}


function range_to_array(results_range) {
  results_range.setBackground("#eeeeee");
  
  var candidates = [];
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      continue;
    }
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      var cell_value = cell.getValue();
      cell.setBackground("#ffff00");
      if (!include(candidates, cell_value)) {
        candidates.push(cell_value);
      }
    }
  }
  return candidates;
}


function get_all_candidates(results_range) {
  results_range.setBackground("#eeeeee");
  
  var candidates = [];
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      continue;
    }
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      var cell_value = cell.getValue();
      cell.setBackground("#ffff00");
      if (!include(candidates, cell_value)) {
        candidates.push(cell_value);
      }
    }
  }
  return candidates;
}


function get_votes(results_range, candidates, keys_range, valid_keys) {
  if (typeof keys_range === "undefined") {
    keys_range = null;
  }
  var votes = {};
  var keys_used = [];
  
  for (var c = 0; c < candidates.length; c++) {
    votes[candidates[c]] = 0;
  }
  
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      break;
    }
    
    if (keys_range != null) {
      // Only use key once.
      var key_cell = keys_range.getCell(row, 1);
      var key_cell_value = key_cell.getValue();
      if (!include(valid_keys, key_cell_value) ||
          include(keys_used, key_cell_value)) {
        key_cell.setBackground('#ffaaaa');
        continue;
      } else {
        key_cell.setBackground('#aaffaa');
        keys_used.push(key_cell_value);
      }
    }
    
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      
      var cell_value = cell.getValue();
      if (include(candidates, cell_value)) {
        votes[cell_value] += 1;
        cell.setBackground("#aaffaa");
        break;
      }
      cell.setBackground("#aaaaaa");
    }
  }
  if (keys_range != null) {
    update_keys_used(keys_used);
  }
  return votes;
}


function update_keys_used(keys_used) {
  // Check to make sure sheet exists //
  var keys_used_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME);
  if (keys_used_sheet == null) {
    missing_keys_used_sheet_alert = true;
    return;
  }
  
  // Clear Keys Used //
  clear_keys_used(); // relies on global VOTE_TYPE_NAME
  
  // Update List // 
  var entryIndex = GetVoteIndex(VOTE_TYPE_NAME)-1;
  var usedKeysRangeString = String.fromCharCode(65+entryIndex) + '2:' + String.fromCharCode(65+entryIndex);
  var usedKeysRange = keys_used_sheet.getRange(usedKeysRangeString);
  for (var row = 0; row < keys_used.length; row++) {
    usedKeysRange.getCell(row+1,1).setValue(keys_used[row]).setBackground('#eeeeee');
  }
}

function clear_keys_used() {
  // Check to make sure sheet exists //
  var keys_used_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME);
  if (keys_used_sheet == null) {
    missing_keys_used_sheet_alert = true;
    return;
  }
  
  // Clear List //
  var entryIndex = GetVoteIndex(VOTE_TYPE_NAME)-1;
  var usedKeysRangeString = String.fromCharCode(65+entryIndex) + '2:' + String.fromCharCode(65+entryIndex);
  var usedKeysRange = keys_used_sheet.getRange(usedKeysRangeString);
  for (var row=0; row<usedKeysRange.getNumRows(); row++)
    usedKeysRange.getCell(row+1,1).setValue("").setBackground('#ffffff');
    
}

function get_winner(votes, candidates) {
  var total = 0;
  var winning = null;
  var max = 0;
  for (var c = 0; c < candidates.length; c++) {
    var name = candidates[c];
    var count = votes[name];
    total += count;
    if (count > max) {
      winning = name;
      max = count;
    }
  }
  
  if (max * 2 > total) {
    return winning;
  }
  return null;
}


function get_remaining_candidates(votes, candidates) {
  var min = -1;
  for (var c = 0; c < candidates.length; c++) {
    var name = candidates[c];
    var count = votes[name];
    if (count < min || min == -1) {
      min = count;
    }
  }
  
  var c = 0;
  while (c < candidates.length) {
    var name = candidates[c];
    var count = votes[name];
    if (count == min) {
      candidates.splice(c, 1);
    } else {
      c++;
    }
  }
  return candidates;
}
  
/*
http://stackoverflow.com/questions/143847/best-way-to-find-an-item-in-a-javascript-array
*/
function include(arr,obj) {
    return (arr.indexOf(obj) != -1);
}


/*
Returns the number of consecutive rows that do not have blank values in the first column.
http://stackoverflow.com/questions/4169914/selecting-the-last-value-of-a-column
*/
function get_num_rows_with_values(results_range) {
  var num_rows_with_votes = 0;
  var num_rows = results_range.getNumRows();
  for (var row = 1; row <= num_rows; row++) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      break;
    }
    num_rows_with_votes += 1;
  }
  return num_rows_with_votes;
}

function ResetBackgroundColors() {
  for(var i=0; i<VOTE_TYPE_COUNT; i++)
  {
    // Clears Highlighting //
    var numColumns = VOTE_TYPE_ARRAY[i].ChoiceCount;
    var baseColumns = VOTE_TYPE_ARRAY[i].BaseColumn;
    clear_background_color(baseColumns, numColumns);
  }
}

function clear_background_color(baseCol, numCols) {
  var results_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, baseCol, numCols);
  if (results_range == null) {
    return;
  }
  results_range.setBackground('#eeeeee');
  
  if (USING_KEYS) {
    var keys_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, KEYS_COLUMN, 1);
    keys_range.setBackground('#eeeeee');
  }
}