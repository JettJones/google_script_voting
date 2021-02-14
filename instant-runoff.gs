/*
Instant Runoff Voting with Google Form and Google Apps Script
Author: Jett Jones

This is a twice forked project, from Darrell Ross, and the original script, written by Chris Cartland.

This project may contain bugs. Use at your own risk.

This script can be configured with the 'Configure' tab in the Spreadsheet.
There should be no need to edit this file at all.

Instructions: https://github.com/JettJones/google_script_voting

*/

/* BEGIN SETTINGS */
/**************************/
const CONFIGURE_SHEET_NAME = "Configure";          // Configuration for the voting process is in this sheet.
const CREDENTIAL_SHEET_NAME = "Credentials";       // Credentials and voter contact information.
const FILTER_SHEET_NAME = "FilteredVotes";         // Votes are validated for Valid keys, then placed in this sheet.
const RESULTS_SHEET_NAME = "Results";              // Logs of voting rounds are placed in this sheet.
const RANDOM_RESULT = "Lots";
const NO_ACTION_KEYS = ["No Action"];
var VOTE_SHEET_NAME = "Votes";                     // Active sheet where form data is entered.
const BASE_ROW = 2;                                // row where votes start - could change it if you wanted to keep old votes
/**************************/
/* END SETTINGS */

function VoteType(pvote,pchoices,pbasecol,presultrow,pcandidates)
{
  this.VoteName = pvote;
  this.ChoiceCount = pchoices;
  this.BaseColumn = pbasecol;
  this.ResultRow = presultrow;
  this.Candidates = pcandidates;
}
/**************************/
var CONFIG;                   // Map of configuration key-values
var USING_KEYS;               // If true, use only votes with valid keys
var SHEET_PER_ROUND;          // Show a separate sheet for each round
var USING_NO_ACTION;          // Enable special handling for 'No Action'
var TOP_N;                    // Stop counting once there are fewer than N candidates
function InitConfig() {
  if (CONFIG == null) {
    ensure_sheet(CONFIGURE_SHEET_NAME, SetupConfigureSheet);
    CONFIG = read_config();
    var usingKeys = CONFIG.get("Use Keys");
    USING_KEYS = is_true(usingKeys.toString());
    var sheetPerRound = CONFIG.get("Sheet Per Round") || "False";
    SHEET_PER_ROUND = is_true(sheetPerRound.toString());
    var useNoAction = CONFIG.get("Use No Action") || "False";
    USING_NO_ACTION = is_true(useNoAction.toString());

    var useTopN = CONFIG.get("Top N") || "0"
    TOP_N = parseInt(useTopN);
  }
}

function is_true(value) {
  return ["yes", "y", "true", "1" ].includes(value.toLowerCase())
}

/**************************/
var VOTE_TYPE_ARRAY = [];
var LOG_LENGTH = null;                       // Row for logs in the result sheet
/**************************/

function read_config() {
  var config_range = get_range_with_values(CONFIGURE_SHEET_NAME, 1, 1, 2);

  if (config_range == null) {
    Browser.msgBox("Configuration missing. Looking within sheet: " + CONFIGURE_SHEET_NAME);
    return false;
  }

  var result = new Map();
  result.set("Vote", new Set());
  const maxRows = config_range.getNumRows();
  for (var row = 1; row <= maxRows; row++) {
    var key = config_range.getCell(row, 1);
    var value = config_range.getCell(row, 2);

    if (key.isBlank()) {
      continue;
    }

    const key_str = key.getValue();
    if (key_str == "Vote") {
      result.get(key_str).add(value.getValue());
    } else {
      result.set(key.getValue(), value.getValue());
    }
  }

  return result;
}

function InitVoteTypesArray() {
  // read the configuration sheet //
  InitConfig();
  var vote_set = CONFIG.get("Vote");

  // assume that the first sheet is the voting sheet
  var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  VOTE_SHEET_NAME = vote_sheet.getSheetName();
  const vote_header_ix  = Math.max(1, vote_sheet.getFrozenRows());
  const max_cols = vote_sheet.getMaxColumns();
  var header_row = vote_sheet.getRange(vote_header_ix, 1, 1, max_cols);
  var headers = header_row.getValues()[0];

  function string_to_name(s) {
      var ixStart = s.lastIndexOf("[");
      var ixEnd = s.lastIndexOf("]");
      var prefix = s.substr(0, ixStart).trim();
      return [prefix, s.substr(ixStart + 1, ixEnd - ixStart - 1)];
  }

  // read through the vote info headers.
  var current_match = null;
  var current_names = [];
  var base_col = -1;
  var base_result = 1;

  function maybe_add() {
    if (current_names.length > 0) {
      if (vote_set.has(current_match)) {
        VOTE_TYPE_ARRAY.push(new VoteType(current_match, current_names.length, base_col, base_result, current_names));
        base_result += current_names.length
      } else {
        Logger.log("Skipping unexpected vote named:" + current_match);
      }
    }
  };

  for (var col = 0; col < headers.length; col++) {
    var name = headers[col];
    if (name === "" || name.indexOf("[") == -1) {
      continue;
    }
    var split = string_to_name(headers[col]);
    if (split[0] != current_match) {
      // this is a new vote name, report the previous if any.
      maybe_add();

      // reset
      [base_col, current_match, current_names] = [col, split[0], []];
    }
    current_names.push(split[1]);
  }

  maybe_add();

  return true;
}

function get_colors(count) {
   const light = [
     // "#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3", "#d0e0e3", "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc"
      "#dd7e6b", "#ea9999", "#f9cb9c", "#ffe599", "#b6d7a8", "#a2c4c9", "#a4c2f4", "#9fc5e8", "#b4a7d6", "#d5a6bd",
      "#cc4125", "#e06666", "#f6b26b", "#ffd966", "#93c47d", "#76a5af", "#6d9eeb", "#6fa8dc", "#8e7cc3", "#c27ba0",
      "#a61c00", "#cc0000", "#e69138", "#f1c232", "#6aa84f", "#45818e", "#3c78d8", "#3d85c6", "#674ea7", "#a64d79",
      "#85200c", "#990000", "#b45f06", "#bf9000", "#38761d", "#134f5c", "#1155cc", "#0b5394", "#351c75", "#741b47",
      //"#5b0f00", "#660000", "#783f04", "#7f6000", "#274e13", "#0c343d", "#1c4587", "#073763", "#20124d", "#4c1130"
   ];

  result = [];
  while(count > 0) {
    if (count > light.length) {
      result.push(light);
      count -= light.length;
    } else {
      result.push(...light.slice(0, count));
      count = 0;
    }
  }

  return result;
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

function initialize_spreadsheet() {
  InitVoteTypesArray();
  create_menu_items();
}

// Clears away old tallied votes. Does not remove configuration.
function clear_voting() {
  initialize_spreadsheet();
  SetupResultsSheet();
  SetupFilteredVotesSheet();
  clear_rounds();
  clear_vote_highlights();
}

function filter_only() {
  initialize_spreadsheet();
  SetupFilteredVotesSheet();
  clear_vote_highlights();
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

  InitConfig();

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
    Browser.msgBox("Vote Tallying complete - see Results tab.");
  else
    Browser.msgBox("Vote Tallying failed.");
}

// Takes votes from the voting sheet, skips duplicates or invalid keys and copies to Filtered
function filter_votes() {
  InitConfig();

  var filter_sheet = ensure_sheet(FILTER_SHEET_NAME);
  var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET_NAME);

  // set the header
  var header = [];
  header.push("Timestamp");

  for(const vote_detail of VOTE_TYPE_ARRAY) {
    while ((header.length) < vote_detail.BaseColumn) {
      header.push("");
    }
    const range_header = Array.from({length: vote_detail.ChoiceCount}, (_, i) => i+1);
    range_header.forEach(function(v) {header.push(v);});
  }

  filter_sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");


  // converting votes from the form output to countable rows
  const other_cols = CONFIG.get("Other Columns") || 0;
  const key_col = CONFIG.get("Key Column") || 2;
  var maxRow = vote_sheet.getLastRow();
  var maxCol = vote_sheet.getLastColumn() - other_cols;
  var full_range = vote_sheet.getDataRange();
  var full_values = full_range.getValues();

  // Rather than highlight and copy rows one-by-one
  // work in batches, for the slow spreadsheet apis.
  var chunk = [];
  var write_ix = 2;
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

  // allow a header on the voting sheet.
  var read_min = Math.max(BASE_ROW, vote_sheet.getFrozenRows() + 1);

  for (var i=maxRow; i >= read_min; i--) {
    const rowIx = i - 1;
    const cell = full_values[rowIx][0];
    if (cell === "") {
      continue;
    }

    if (USING_KEYS) {
      const keyVal = full_values[rowIx][key_col];
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
      const colIx = vote_detail.BaseColumn;
      const endColIx = colIx + vote_detail.ChoiceCount;
      const votes = full_values[rowIx].slice(colIx, endColIx);

      //
      // Votes come in the format 1st 2nd 3rd ... etc
      // First map each vote to a pair with candidate name, like: ["5th", "Jo Joeson"]
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
    write_chunk(read_min);
  }

  if (USING_KEYS) {
    // D represents the last column before votes, this could be the lowest BaseColumn in vote_detail above
    const cells = errorRows.map(([ix, _]) => "A" + ix + ":D" + ix);
    vote_sheet.getRangeList(cells).setBackground("#FF9900");
    for (const [ix, note] of errorRows) {
      vote_sheet.getRange(ix, 2).setNote(note);
    }
  }
}

function get_valid_keys() {
  var range = get_range_with_values(CREDENTIAL_SHEET_NAME, base_row=2, base_column=3, num_columns=1)
  if (range) {
    return range.getValues().map(([x]) => x);
  } else {
    return [];
  }
}

// Creates the Configure Sheet if it doesn't exist. Does not remove configuration data.
function SetupConfigureSheet() {
  var config = ensure_sheet(CONFIGURE_SHEET_NAME);

  const key_setting = config.getRange("B1").getValue();
  const header = [["Use Keys"],["Vote"]];
  const notes = [["Set  cell B1 to 'Yes' to enable keys."],
      ["Enter the name of the vote you are holding."]];

  config.getRange("A1:A2").setValues(header).setNotes(notes).setFontWeight("bold");
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

  result_sheet.getRange(1, 1, 1, 5).setValues([["Vote Name", "Rounds", "", "", "Log Length:"]]).setFontStyle("bold");

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

function results_for_round(vote_detail, roundN, votes) {
  var result_sheet = ensure_sheet(RESULTS_SHEET_NAME);
  var row_start = vote_detail.ResultRow + 1; // where to start writing the election output

  var range = result_sheet.getRange(row_start, roundN + 1, vote_detail.Candidates.length + 1, 1);

  var vote_counts = [[roundN]];
  vote_counts.push(...vote_detail.Candidates.map( (name) => [(votes.get(name) || [] ).length]));

  range.setValues(vote_counts);
}

function tally_single_vote(voteType) {
  /* Determine number of voting columns */
  var choiceColumnCount = voteType.ChoiceCount;
  var baseColumn = voteType.BaseColumn + 1;
  var baseRow = 2
  var round = 1;

  /* Begin */
  var input_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FILTER_SHEET_NAME);
  var results_range = get_range_with_values(FILTER_SHEET_NAME, baseRow, baseColumn, choiceColumnCount);

  if (results_range == null) {
    Browser.msgBox("No votes. Looking for sheet: " + FILTER_SHEET_NAME);
    return false;
  }

  /* candidates is a set of names (strings) */
  var candidates = get_all_candidates(results_range);
  const colors = get_colors(candidates.size);
  var color_map = new Map();
  var ix=0;
  candidates.forEach((c) => color_map.set(c, colors[ix++]));

  // Coloring cells one by one is slow
  // Instead, collect cells by the intended color, and setBackground in bulk
  var color_updates = new Map();
  const show_color = function(color_map) {
    for(const [color, list] of color_map) {
      const r1c1 = list.map(function([r, c]) {return String.fromCharCode(63 + c + baseColumn) + (r + 1);});
      input_sheet.getRangeList(r1c1).setBackground(color);
    }
  };

  /* votes is an map from candidate names -> list of votes */
  var votes = get_votes(results_range, candidates, color_updates, color_map);

  /* winner is candidate name (string) or null */
  var winner = get_winner(votes, candidates);

  var status = [];
  show_color(color_updates);
  status.push(summary_string(votes, winner));
  results_for_round(voteType, round, votes);

  function copy_round(r) {
    if (!SHEET_PER_ROUND) { return; }

    const round_name = "_round" + r;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(round_name);
    if (sheet != null) { ss.deleteSheet(sheet); }
    const len = ss.getSheets().length
    ss.insertSheet(round_name, sheet_index=len, options={'template':input_sheet});
  }

  copy_round(round);
  round += 1;

  while (winner == null) {

    if (TOP_N > 0 && candidates.size <= TOP_N) {
      status.push("Found Top " + TOP_N + " candidates")
      show_status(status);
      show_status(["Result: Complete"]);
      return true;
    }

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
      show_status(["Result: TIE "]);
      return true;
    }

    status = [];
    color_updates.clear();
    votes = update_votes(results_range, candidates, votes, eliminated, color_updates, color_map);
    winner = get_winner(votes, candidates);

    show_color(color_updates);
    status.push(summary_string(votes, winner))
    results_for_round(voteType, round, votes);
    copy_round(round);
    round += 1;
  }

  show_status(status);
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

function get_votes(results_range, candidates, color_updates, color_map) {
  var votes = new Map();

  for(const c of candidates) {
    votes.set(c, []);
  }

  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    row_vote(results_range, candidates, votes, row, num_columns, color_updates, color_map);
  }
  return votes;
}

function row_vote(results_range, candidates, votes, row, num_columns, color_updates, color_map) {
  function set_color(value, color) {
    if (!color_updates.has(color)) { color_updates.set(color, []); }
    color_updates.get(color).push(value);
  }

    for (var column = 1; column <= num_columns; column++) {
      const cell = results_range.getCell(row, column);
      const is_blank = cell.isBlank();
      var cell_value;
      if (USING_NO_ACTION && cell.isBlank()) { cell_value = NO_ACTION_KEYS[0]; }
      else if (is_blank) { break; } // no futher votes for this row; exhausted
      else { cell_value = cell.getValue(); }

      if (candidates.has(cell_value)) {
        votes.get(cell_value).push(row);
        const color = (color_map) ? color_map.get(cell_value) : "#aaffaa";
        set_color([row, column], color);
        break;
      }

      set_color([row, column], "#aaaaaa");
    }
}

function update_votes(results_range, candidates, votes, removed, color_updates, color_map) {
  var num_columns = results_range.getNumColumns();
  for (const name of removed) {
    const rows = votes.get(name);
    for(const row of rows) {
      row_vote(results_range, candidates, votes, row, num_columns, color_updates, color_map);
    }
    votes.delete(name);
  }
  return votes;
}

function get_winner(votes) {
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
    if (USING_NO_ACTION && NO_ACTION_KEYS.includes(name)) { continue;}

    var count = votes.get(name).length;
    if (count < min || min == -1) {
      min = count;
    }
  }

  var remove = [];
  for(const name of candidates) {
    if (USING_NO_ACTION && NO_ACTION_KEYS.includes(name)) { continue;}

    var count = votes.get(name).length;
    if (count == min)
      remove.push(name);
  }

  return remove;
}

function elimination_tiebreak(results_range, votes, candidates, removed) {
  var result = eliminate_by_next(results_range, votes, candidates, removed);

  if (result.length > 1) {
    result = eliminate_by_lot(result);
  }
  return result;
}

function eliminate_by_lot(removed) {
  // randomly choose one of the removed candidates, and preserve the random choice

  var lot_sheet = ensure_sheet(RANDOM_RESULT);
  var lot_range = get_range_with_values(RANDOM_RESULT, 1, 1, 2);

  var result = new Map();
  const maxRows = lot_range.getNumRows();
  for (var row = 1; row <= maxRows; row++) {
    var key = lot_range.getCell(row, 1);

    if (key.isBlank()) {
       continue;
    }

    var value = lot_range.getCell(row, 2);
    result.set(key.getValue(), value.getValue());
  }

  var removed_list = new Array(removed);
  removed_list.sort();
  var lot_key = removed_list.join("||");

  var chosen = 0;
  if (result.has(lot_key)) {
    chosen = result.get(lot_key);
  } else {
    chosen = Math.floor(Math.random() * removed.length);
    lot_sheet.getRange(lot_range.getNumRows() + 1, 1, 1, 2).setValues([[lot_key, chosen]]);
  }
  return [removed[chosen]];
}

function eliminate_by_next(results_range, votes, candidates, removed) {
  // for all removed candidates, check the votes in the next round
  // the candidate with the least follow-on votes is removed.

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
      row_vote(results_range, temp_candidates, next_votes, row, num_columns, _ignore_colors, null);
    }
  }

  return get_eliminated_candidates(next_votes, removed);
}

/**************************/
 function make_test_votes() {
  var count = 100;

  InitVoteTypesArray();
  var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET_NAME);
  const last_row = vote_sheet.getLastRow();
  const vote = VOTE_TYPE_ARRAY[0];

  function gen_vote() {
    var row = [];
    row[0] = new Date().toLocaleString("en-US");
    row[1] = "TEST GENERATED";

    ranks = Array.from({length: vote.ChoiceCount}, (_, i) => i+1);
    ranks.sort((a,b) => 0.5 - Math.random());

    for(var i = 0; i < vote.ChoiceCount; i++) {
      row[i + vote.BaseColumn] =  ranks[i];
    }
    return row;
  }
  for (var ix = 0; ix < count; ix++) {
    row = gen_vote();
    var row_ix = last_row + 1 + ix;
    var range = vote_sheet.getRange(row_ix, 1, 1, row.length);
    range.setValues([row]);
  }
}


/**************************/
/**************************/
function clear_rounds() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  for (var ss of sheet.getSheets()) {
    if (ss.getSheetName().startsWith("_round")) {
      sheet.deleteSheet(ss);
    }
  }
}

function clear_vote_highlights() {
  var vote_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VOTE_SHEET_NAME);
  const other_cols = CONFIG.get("Other Columns") || 0;
  var maxRow = vote_sheet.getLastRow();
  var maxCol = vote_sheet.getLastColumn() - other_cols;

  var range = vote_sheet.getRange(1, 1, maxRow, maxCol);
  range.clearFormat();

}

function ensure_sheet(sheet_name, callback) {
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (active_spreadsheet.getSheetByName(sheet_name) == null) {
    const ix = active_spreadsheet.getSheets().length;
    active_spreadsheet.insertSheet(sheet_name, sheetIndex=ix);
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
  if (max_row > 0)
    results_range = results_sheet.getRange(base_row, base_column, max_row - base_row + 1, num_columns);
  else
    results_range = results_sheet.getDataRange()
  return results_range;
}

function clear_background_color(sheet_name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  sheet.clearFormats().clearNotes();
}
