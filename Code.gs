/*
Instant-runoff voting with Google Form and Google Apps Script
Author: Zoe Smith
Date created: 2021-04-27
Last code update: 2021-04-30

STILL IN PROGRESS

Based on code from Chris Cartland: https://github.com/cartland/instant-runoff

This code is set up to have two ballots in the same form.

Steps to run an election.
* Create a new Google Form.
* From the form spreadsheet go to "Tools" -> "Script Editor..."
* Copy the code into the editor.
* Configure settings in the editor and match the settings with the names of your sheets.
* Run create_menu_items() from the Script Editor to get menu options in spreadsheet
* Send out the live form for voting
* From the form spreadsheet go to "Instant Runoff" -> "Run".
    * If this is not an option, run the function run_instant_runoff() directly from the Script Editor.
*/


/* Settings */ 

var VOTE_SHEET_NAME = "Form Responses 1";
var BASE_ROW = 2;
var BASE_COLUMN = 2;
var NUM_COLUMNS = 3;

/* End Settings */

function create_menu_items() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [ {name: "Run", functionName: "run_instant_runoff"},
                        {name: "Reset Colors", functionName: "clear_background_color"} ];
    ss.addMenu("Instant Runoff", menuEntries);
}

/* Create menus */
function onOpen() {
    create_menu_items();
}

/* Create menus when installed */
function onInstall() {
    onOpen();
}

function run_instant_runoff() {
  /* Begin */
  clear_background_color();
  
  var results_range1 = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, BASE_COLUMN, NUM_COLUMNS);
  var winner1;
  
  if (results_range1 == null) {
    winner1 = "No votes for full-time.";
  }
  winner1 = run_ballot(results_range1);

  var results_range2 = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, BASE_COLUMN+NUM_COLUMNS, NUM_COLUMNS);
  var winner2;

  if (results_range2 == null) {
    winner2 = "No votes for part-time.";
  }
  winner2 = run_ballot(results_range2);
  
  var winner1_message = "Full-Time Winner: " + winner1;
  var winner2_message = "Part-Time Winner: " + winner2;
  Browser.msgBox(winner1_message+'\\n'+winner2_message);
}

// Returns the name of the winner
// Can I make this give me the vote percentage too?
function run_ballot(results_range) {
  /* candidates is a list of names (strings) */
  var candidates = get_all_candidates(results_range);
  
  /* votes is an object mapping candidate names -> number of votes */
  var votes = get_votes(results_range, candidates);
  
  /* winner is candidate name (string) or null */
  var winner = get_winner(votes, candidates);

  while (winner == null) {
    /* Modify candidates to only include remaining candidates */
    get_remaining_candidates(votes, candidates);
    if (candidates.length == 0) {
      winner = "Tie";
    }
    votes = get_votes(results_range, candidates);
    winner = get_winner(votes, candidates);
  }
  return winner;
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


function get_votes(results_range, candidates) {
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
  return votes;
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

// Returns the number of consecutive rows that do not have blank values in the first column.
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

// Returns the number of consecutive columns that do not have blank values in the first row.
function get_num_columns_with_values(results_range) {
  var num_columns_with_values = 0;
  var num_columns = results_range.getNumColumns();
  for (var col = 1; col <= num_columns; col++) {
    var first_is_blank = results_range.getCell(1, col).isBlank();
    if (first_is_blank) {
      break;
    }
    num_columns_with_values += 1;
  }
  return num_columns_with_values;
}

function clear_background_color() {
  var results_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, BASE_COLUMN, NUM_COLUMNS*2);
  if (results_range == null) {
    return;
  }
  results_range.setBackground('#ffffff'); // set background to white
}
