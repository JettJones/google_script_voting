/*
Instant Runoff Voting with Google Form and Google Apps Script
Author: Jett Jones

This is a twice forked project, from Darrell Ross, and the original script, written by Chris Cartland.

This project may contain bugs. Use at your own risk.

This script can be configured with the 'Configure' tab in the Spreadsheet.
There should be no need to edit this file at all.

Instructions: https://docs.google.com/a/milsoft.com/document/d/1bLEOPwxfSTwEh_pQglM7nGkUtaFkjGu5QWqEV17FhcE/edit?usp=sharing

*/

/* BEGIN SETTINGS */
/**************************/
var VOTE_SHEET_NAME = "Votes";                   // active sheet where form data is entered will be renamed to this
var CONFIGURE_SHEET_NAME = "Configure";          // Configuration for the voting process and secret keys is in this sheet.
var RESULTS_SHEET_NAME = "Results";              // Logs of voting rounds are placed in this sheet.
var FILTER_SHEET_NAME = "FilteredVotes";         // Votes are validated for Valid keys, then placed in this sheet.
var BASE_ROW = 2;                                // row where votes start - could change it if you wanted to keep old votes
/**************************/
/* END SETTINGS */

function VoteType(pvote,pindex,pchoices,pbasecol,pcandidates)
{
  this.VoteName = pvote;
  this.VoteIndex = pindex;
  this.ChoiceCount = pchoices;
  this.BaseColumn = pbasecol;
  this.Candidates = pcandidates;
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
  var config_sheet = ensure_sheet(CONFIGURE_SHEET_NAME, SetupConfigureSheet);
  var usingKeys = config_sheet.getRange("B1").getValue();
  USING_KEYS = (usingKeys.toString().toLowerCase() == "yes")
}

/**************************/
var VOTE_TYPE_COUNT = 0;                     // Number of vote types - set during InitVoteTypesArray();
var VOTE_TYPE_ARRAY = [];
var WINNER_ARRAY = [];                       // All winner data
var LOG_LENGTH = null;                       // Row for logs in the result sheet
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

  var previousChoiceCount = 0;
  const maxRows = results_range.getNumRows();

  function get_candidate_names(base_column, choice_count) {
    var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET_NAME);
    var range = vote_sheet.getRange(1,  base_column, 1, choice_count);

    const string_to_name = function (s) {
      var ixStart = s.lastIndexOf("[");
      var ixEnd = s.lastIndexOf("]");
      return s.substr(ixStart + 1, ixEnd - ixStart - 1);
    }

    return range.getValues()[0].map(string_to_name);
  }

  // read configuration for each vote type
  for (var row = 1; row <= maxRows; row++) {
    var voteTypeCell = results_range.getCell(row, 1);    // Vote Type is First Column of results_range
    if (voteTypeCell.isBlank()) {
      continue;
    }

    var voteTypeCellValue = voteTypeCell.getValue();
    var choiceCountCell = results_range.getCell(row, 2); // Vote Choice Count is Second Column of results_range
    var choiceCountCellValue = choiceCountCell.getValue();
    baseCol += previousChoiceCount;
    previousChoiceCount = choiceCountCellValue;
    var candidateNames = get_candidate_names(baseCol, choiceCountCellValue);
    VOTE_TYPE_ARRAY.push(new VoteType(voteTypeCellValue, row, choiceCountCellValue, baseCol, candidateNames));
  }

  VOTE_TYPE_COUNT = VOTE_TYPE_ARRAY.length;
  return true;
}


/**************************/

/***** MENU CONFIGURATION *****/
function create_menu_items() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [ {name: "Initialize Spreadsheet", functionName: "initialize_spreadsheet"},
                        {name: "Clear Tally", functionName: "clear_voting"},
                        {name: "Prepare Votes", functionName: "filter_only"},
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

/**************************/
/**************************/
/**************************/

function OutputVoteResults() {
  var fullResultString = "";
  for (var i=0; i<VOTE_TYPE_COUNT; i++) {
    var winner = WINNER_ARRAY[i].WinnerName;
    var winnerVote = WINNER_ARRAY[i].WinnerVote;
    var winnerDate = WINNER_ARRAY[i].WinnerDate;
    var winnerMessage = "Winner: " + winner + winnerDate;
    fullResultString += "** " + winnerVote + " " + winnerMessage + " **";
  }
  Browser.msgBox(fullResultString);
}

function initialize_spreadsheet() {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  active_spreadsheet.getSheets()[0].setName(VOTE_SHEET_NAME); // Rename Form Entries Sheet to "Votes"
  InitVoteTypesArray();
  create_menu_items();
}

// Clears away old tallied votes. Does not remove configuration.
function clear_voting() {
  initialize_spreadsheet();
  SetupResultsSheet();
  SetupFilteredVotesSheet();
  clear_background_color(VOTE_SHEET_NAME);
}

function filter_only() {
  initialize_spreadsheet();
  SetupFilteredVotesSheet();
  clear_background_color(VOTE_SHEET_NAME);
  filter_votes();
}

function test_tally() {
  InitVoteTypesArray();
  LOG_LENGTH = get_log_length();
  tally_single_vote(VOTE_TYPE_ARRAY[0]);
}

// The full pipeline of tallying votes from form results
function tally_votes() {
  Logger.log("Running Init");
  clear_voting(); // start clean

  InitUsingKeys();

  Logger.log("Running Filter Step");
  filter_votes();

  Logger.log("Running tally");
  var success = true;
  for (const vote_detail of VOTE_TYPE_ARRAY) {
    if(!tally_single_vote(vote_detail)) {
      success = false;
      break;
    }
  }

  Logger.log("Displaying output message");
  if(success)
    OutputVoteResults();
  else
    Browser.msgBox("Vote Tallying failed.");
}

// Takes votes from the voting sheet, skips duplicates or invalid keys and copies to Filtered
function filter_votes() {
  var filter_sheet = ensure_sheet(FILTER_SHEET_NAME);
  var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET_NAME);

  // set the header
  var header = [];
  header.push("Timestamp");

  for(const vote_detail of VOTE_TYPE_ARRAY) {
    while ((header.length + 1) < vote_detail.BaseColumn) {
      header.push("");
    }
    const range_header = Array.from({length: vote_detail.ChoiceCount}, (_, i) => i+1);
    range_header.forEach(function(v) {header.push(v);});
  }

  filter_sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");


  // converting votes from the form output to countable rows
  var maxRow = vote_sheet.getLastRow();
  var maxCol = vote_sheet.getLastColumn();
  var full_range = vote_sheet.getDataRange();
  var full_values = full_range.getValues();

  // Rather than highlight and copy rows one-by-one
  // work in batches, for the slow spreadsheet apis.
  var chunk = [];
  var write_ix = BASE_ROW;
  var last_write = maxRow;
  const max_chunk = 5;

  function write_chunk(row) {
      // add data to filter_sheet
      var write_range = filter_sheet.getRange(write_ix, 1, chunk.length, maxCol);
      write_range.setValues(chunk);
      write_ix += chunk.length;

      // highlight counted rows in vote_sheet
      vote_sheet.getRange(row, 1, last_write - row + 1, maxCol).setBackground("#FFFF77");
      last_write = row;
      chunk = [];
  }

  var seenKeys = new Set();
  var validKeys = new Set(get_valid_keys());
  var errorRows = [];

  for (var i=maxRow; i >= BASE_ROW; i--) {
    const rowIx = i - 1;
    const cell = full_values[rowIx][0];
    if (cell === "") {
      continue;
    }

    if (USING_KEYS) {
      const keyVal = full_values[rowIx][1];
      if (keyVal === "") {
        errorRows.push([i, "Key was not set"]);
        continue;
      }
      if (seenKeys.has(keyVal)) {
        errorRows.push([i, "Key Reused Later"]);
        continue;
      }
      if (!validKeys.has(keyVal)) {
        errorRows.push([i, "Key was not valid"]);
        continue;
      }
      seenKeys.add(keyVal);
    }

    // Processing for a valid vote
    var row_values = new Array(maxCol);
    chunk.push(row_values);
    row_values[0] = cell;

    for(const vote_detail of VOTE_TYPE_ARRAY) {
      const candidate_names = vote_detail.Candidates;

      // convert from row, column to zero-based indexes in full_values
      const colIx = vote_detail.BaseColumn -1;
      const endColIx = colIx + vote_detail.ChoiceCount;
      const votes = full_values[rowIx].slice(colIx, endColIx);

      //
      // Votes come in the format 1st 2nd 3rd ... etc
      // First map each vote to a pair with candidate name, like: ["5th", "Joe Joeson"]
      // Drop columns with a blank / missing vote.
      // Then parse the order so "12th" becomes 12.
      // Then sort pairs by that order.
      // Then keep only the candidate names.
      //
      var vote_pair = votes.map( function(val, ix) { return [val, candidate_names[ix]];} );
      vote_pair = vote_pair.filter(function (ar) { return (parseInt(ar[0]) > 0);});
      vote_pair.sort(function(a, b){ return parseInt(a[0]) - parseInt(b[0]);});
      var selected = vote_pair.map(function(ar){ return ar[1];});

      row_values.splice(colIx, selected.length, ...selected);
    }

    if (chunk.length >= max_chunk) {
      write_chunk(i);
    }
  }

  if (chunk.length > 0) {
    write_chunk(BASE_ROW);
  }

  if (USING_KEYS) {
    const cells = errorRows.map(([ix, _]) => "A" + ix + ":B" + ix);
    vote_sheet.getRangeList(cells).setBackground("#FF9900");
    for (const [ix, note] of errorRows) {
      vote_sheet.getRange(ix, 2).setNote(note);
    }
  }
}

function get_valid_keys() {
  var range = get_range_with_values(CONFIGURE_SHEET_NAME, 2, 1, 1)
  return range.getValues().map(([x]) => x);
}

// Creates the Configure Sheet if it doesn't exist. Does not remove configuration data.
function SetupConfigureSheet() {
  var config = ensure_sheet(CONFIGURE_SHEET_NAME);

  const key_setting = config.getRange("B1").getValue();
  const header = ["Keys",key_setting, "Votes", "Choice Counts"];
  const notes = ["Enter the keys in this column, starting on the second row. One key per cell.",
  "Set this cell to 'Yes' to enable keys.",
  "Enter the names of each vote you are holding. Enter them in the same order as you have them on your form.",
  "Enter the quantity of choices you have. If this is First Passed The Post voting, then all entries will be 1. If it is Instant-Runoff, enter the number of choices for each item."];

  config.getRange("A1:D1").setValues([header]).setNotes([notes]).setFontWeight("bold");
}

function SetupFilteredVotesSheet() {
  var filter_sheet = ensure_sheet(FILTER_SHEET_NAME);

  filter_sheet.clear({contentsOnly: true, formatOnly: true});
}

// Ensure the RESULTS_SHEET_NAME Sheet Exists and is reset to headers
function SetupResultsSheet() {
  var result_sheet = ensure_sheet(RESULTS_SHEET_NAME);

  // Clear and Repopulate Column Headers //
  result_sheet.clear({commentsOnly: true, contentsOnly: true});

  result_sheet.getRange(1, 1, 1, 5).setValues([["Vote Name", "Winner", "Rounds", "", "Log Length:"]]).setFontStyle("bold");

  var output_row = 2;

  for(const vote_detail of VOTE_TYPE_ARRAY) {
    result_sheet.getRange(output_row, 1).setValue(vote_detail.VoteName).setFontStyle("bold");
    output_row += 1;
    result_sheet.getRange(output_row, 1, vote_detail.Candidates.length, 1).setValues(vote_detail.Candidates.map((name) => [name])).setHorizontalAlignment("right");
    output_row += vote_detail.Candidates.length;
  }

  LOG_LENGTH = output_row;
  result_sheet.getRange(1,6).setValue(LOG_LENGTH);
}

function results_for_round(vote_detail, roundN, votes)
{
  var result_sheet = ensure_sheet(RESULTS_SHEET_NAME);

  var range = result_sheet.getRange(vote_detail.BaseColumn, roundN + 1, vote_detail.Candidates.length, 1);

  var vote_counts = vote_detail.Candidates.map( (name) => [(votes.get(name) || [] ).length]);

  range.setValues(vote_counts);
}

function tally_single_vote(voteType) {
  var voteTypeName = voteType.VoteName;

  /* Determine number of voting columns */
  var choiceColumnCount = voteType.ChoiceCount;
  var baseColumn = voteType.BaseColumn;
  var round = 1;

  /* Begin */
  var input_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FILTER_SHEET_NAME);
  var results_range = get_range_with_values(FILTER_SHEET_NAME, BASE_ROW, baseColumn, choiceColumnCount);

  if (results_range == null) {
    Browser.msgBox("No votes. Looking for sheet: " + FILTER_SHEET_NAME);
    return false;
  }

  /* candidates is a set of names (strings) */
  var candidates = get_all_candidates(results_range);

  // Coloring cells one by one is slow
  // Instead, collect cells by the intended color, and setBackground in bulk
  var color_updates = new Map();
  const show_color = function(color_map) {
    for(const [color, list] of color_map) {
      const r1c1 = list.map(function([r, c]) {return String.fromCharCode(63 + c + baseColumn) + (r + BASE_ROW -1);});
      input_sheet.getRangeList(r1c1).setBackground(color);
    }
  };

  /* votes is an map from candidate names -> list of votes */
  var votes = get_votes(results_range, candidates, color_updates);

  /* winner is candidate name (string) or null */
  var winner = get_winner(votes, candidates);

  var status = [];
  show_color(color_updates);
  status.push(summary_string(votes, winner));
  results_for_round(voteType, round, votes);
  round += 1;

  while (winner == null) {
    var eliminated = get_eliminated_candidates(votes, candidates);

    if (eliminated.length > 1) {
      status.push("Elimination tie - checking second preferences");
      eliminated = elimination_tiebreak(results_range, votes, candidates, eliminated);
    }

    for(const name of eliminated){
      candidates.delete(name);
    }

    status.push(summary_eliminated(eliminated));
    show_status(status);

    if (candidates.size == 0) {
      var dateTimeMessage = " \nDate and time: " + Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
      WINNER_ARRAY.push(new WinnerData("TIE",voteTypeName,dateTimeMessage));
      show_status(["Result: TIE "]);
      return true;
    }

    status = [];
    color_updates.clear();
    votes = update_votes(results_range, candidates, votes, eliminated, color_updates);
    winner = get_winner(votes, candidates);

    show_color(color_updates);
    status.push(summary_string(votes, winner))
    results_for_round(voteType, round, votes);
    round += 1;
  }

  show_status(status);
  var dateTimeMessage = " \nDate and time: " + Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
  WINNER_ARRAY.push(new WinnerData(winner,voteTypeName,dateTimeMessage));

  show_status(["Result: Winner"]);
  return true;
}

function summary_string(votes, winner) {
  var result = "";
  for (const [name, spots] of votes) {
    result += name + " has " + spots.length + ", ";
  }
  if (winner != null) {
    result += winner +  " wins!";
  }
  return result;
}

function summary_eliminated(removed) {
  var result = "Removing lowest votes: " + removed.join(", ");

  return result;
}

function show_status(status_lines) {
  var full_status = status_lines.join("\n");

  var result_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME);

  result_sheet.getRange(LOG_LENGTH, 1).setValue(full_status).setWrap(true);
  LOG_LENGTH = LOG_LENGTH + 1;
  result_sheet.getRange(1,6).setValue(LOG_LENGTH);
}

function get_log_length() {
  var result_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET_NAME);
  return result_sheet.getRange(1,6).getValue() || 2;
}

function get_all_candidates(results_range) {
  results_range.setBackground("#eeeeee");

  var candidates = new Set();
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
      candidates.add(cell_value);
    }
  }
  return candidates;
}

function get_votes(results_range, candidates, color_updates) {
  var votes = new Map();

  for(const c of candidates) {
    votes.set(c, []);
  }

  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    row_vote(results_range, candidates, votes, row, num_columns, color_updates);
  }
  return votes;
}

function row_vote(results_range, candidates, votes, row, num_columns, color_updates) {
  function set_color(value, color) {
    if (!color_updates.has(color)) { color_updates.set(color, []);}
    color_updates.get(color).push(value);
  }

    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        // no further votes for this row; exhausted
        break;
      }

      var cell_value = cell.getValue();
      if (candidates.has(cell_value)) {
        votes.get(cell_value).push(row);
        set_color([row, column], "#aaffaa");
        break;
      }

      set_color([row, column], "#aaaaaa");
    }
}

function update_votes(results_range, candidates, votes, removed, color_updates) {
  var num_columns = results_range.getNumColumns();
  for (const name of removed) {
    const rows = votes.get(name);
    for(const row of rows) {
      row_vote(results_range, candidates, votes, row, num_columns, color_updates);
    }
    votes.delete(name);
  }
  return votes;
}

function get_winner(votes, candidates) {
  var total = 0;
  var winning = null;
  var max = 0;
  for(const [name, spots] of votes) {
    const count = spots.length;
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

function get_eliminated_candidates(votes, candidates) {
  var min = -1;
  for (const name of candidates) {
    var count = votes.get(name).length;
    if (count < min || min == -1) {
      min = count;
    }
  }

  var remove = [];
  for(const name of candidates) {
    var count = votes.get(name).length;
    if (count == min)
      remove.push(name);
  }

  return remove;
}

// Order eliminated candidates by their second preferences
// So if we had Jess and Kat with 3 votes each, and next votes
// went 2 to Kat, 1 to Jess, this method will return Jess to process first
function elimination_tiebreak(results_range, votes, candidates, removed) {
  var next_votes = new Map();
  var _ignore_colors = new Map();

  for(const c of candidates) {
    next_votes.set(c, []);
  }

  var num_columns = results_range.getNumColumns();
  for (const name of removed) {
    const rows = votes.get(name);
    var temp_candidates = new Set(candidates);
    temp_candidates.delete(name);
    for(const row of rows) {
      row_vote(results_range, temp_candidates, next_votes, row, num_columns, _ignore_colors);
    }
  }

  return get_eliminated_candidates(next_votes, removed);
}

function ensure_sheet(sheet_name, callback) {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (active_spreadsheet.getSheetByName(sheet_name) == null) {
    active_spreadsheet.insertSheet(sheet_name);
    if (callback != null) {
      callback();
    }
  }
  return active_spreadsheet.getSheetByName(sheet_name);
}

function get_range_with_values(sheet_string, base_row, base_column, num_columns) {
  var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_string);
  if (results_sheet == null) {
    return null;
  }

  var max_row = results_sheet.getLastRow();
  results_range = results_sheet.getRange(base_row, base_column, max_row - base_row + 1, num_columns);
  return results_range;
}

function clear_background_color(sheet_name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  sheet.clearFormats().clearNotes();
}
