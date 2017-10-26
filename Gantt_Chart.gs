/*
GanttChart_v2.gs
Authors: Jake Bruemmer

This Google Apps Script provides methods and functions for making a simple Gantt Chart
(https://en.wikipedia.org/wiki/Gantt_chart). The idea is for this script to make Gantt
Chart creation very simple in a Google Sheet without taking a bunch of time. For that
reason, there is limited functionality outside of adding, deleting, formatting, and
sorting in the chart.

Naming conventions:
- All functions that being with '_' are meant to be called in other functions
- snake_case because it's easier to read than camelCase
- Functions that don't begin with '_' are called directly by a button in the Sheet
*/

/* ===== Sheet Creation Functions ===== */
/*
Create a new sheet with a Gantt Chart area. Default start and end dates are
currently given. 

Params: none

Returns: none
*/
function create_sheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.insertSheet('Project Plan');
  /* Create all of the headers first */
  /* Then set formulas, named ranges, and hide columns */
  var last_cell = _set_gantt_range(sheet, 43031, 43283);
  _set_days_remaining_values(sheet, last_cell);
  _set_gantt_chart_headers(sheet, last_cell);
  _hide_columns(sheet);
  _format_project_area(sheet);
  _set_frozen_rows_and_cols(sheet);
  format_category_names();
}


/* ===== Sheet Functions ===== */

/*
Gets the Sheet object of the 'Project Plan' sheet.

Params: none

Returns: Sheet
*/
function _get_project_plan_sheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Project Plan');
}

/* ===== Protected Ranges Functions ===== */

/* 
Protects the areas that aren't in the Gantt Chart area for user input. This function is
called in the format_category_names() sheet. It also removes editing ability for all
users other than user that calls the function.

Params: none

Returns: none
*/
function _protect_ranges() {
  /*
  var sheet = _get_project_plan_sheet();
  var gantt_chart = sheet.getRange('gantt_chart');
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
  */
  /* Store the protected ranges in a dictionary for ease of removing editing permissions */
  /*
  var protected_ranges = {};
  protected_ranges['left'] = sheet.getRange(1, gantt_chart.getColumn() - 1, sheet.getMaxRows(), 7);
  protected_ranges['top'] = sheet.getRange(1, gantt_chart.getColumn(), gantt_chart.getRow() - 1, gantt_chart.getNumColumns());
  protected_ranges['right'] = sheet.getRange(1, gantt_chart.getLastColumn() + 1, sheet.getMaxRows(), sheet.getMaxColumns() - gantt_chart.getLastColumn());
  protected_ranges['bottom'] = sheet.getRange(gantt_chart.getLastRow() + 1, gantt_chart.getColumn(), sheet.getMaxRows(), gantt_chart.getNumColumns());
  */
  /* Remove editing privileges */
  /*
  var cur_user = Session.getEffectiveUser();
  for (var key in protected_ranges) {
    var p = protected_ranges[key].protect().setDescription(key + ' of project area.');
    p.addEditor(cur_user);
    p.removeEditors(p.getEditors());
  } */
}

/* ====== Category Functions ====== */
/*
Gets all of the category values in the Task (by Category) column. This function is used
in other functions for formatting, indexing, and copying information for the hidden
sorting columns.

Params: none

Returns: Category and task names {Array}
*/
function _get_category_names() {
  var sheet = _get_project_plan_sheet();
  var gantt_chart = sheet.getRange('gantt_chart');
  return gantt_chart.offset(0, 6, gantt_chart.getNumRows(), 1).getValues();
}

/*
All categories in the simple Gantt Chart are just any pieces of text that don't
start with a hyphen ('-'). Google Apps Script doesn't allow you to use the
Javascript '.startsWith()' method on strings, which is why this function uses the
'.charAt()' method.

Params: 'string' the text to analyze {String}

Returns: True/false of whether text is a category {Boolean}
*/
function _is_category(string) {
  return string.charAt(0) != '-' && string.charAt(0) != '';
}

/*
Gets the indexes of the category names in the project area relative to the top row
of the Gantt Chart. Function is used for formatting the categories.

Params: none

Returns: Array of indexes of the categories {Array}
*/
function _get_indexes_for_categories() {
  var categories = _get_category_names();
  var indexes = [];
  for (var i = 0; i < categories.length; i++) {
    if (_is_category(String(categories[i]))) indexes.push(i);
  }
  return indexes;
}

/*
Copies the category name for all of the categories and tasks to column B, which is
hidden in the sheet. Column B is used for sorting the project area if the user wants
to sort the project area.

Params: none

Returns: none
*/
function _write_category_names() {
  var indexes = _get_indexes_for_categories();
  var categories = _get_category_names();
  var sheet = _get_project_plan_sheet();
  var gantt_chart = sheet.getRange('gantt_chart');
  for (var i = 0; i < indexes.length; i++) {
    for (var j = indexes[i]; j < indexes[i + 1]; j++) {
      gantt_chart.getCell(j + 1, 1).setValue(categories[indexes[i]]);
    }
  }
  for (var i = indexes[indexes.length - 1]; i < gantt_chart.getNumRows() - 1; i++) {
    gantt_chart.getCell(i + 1, 1).setValue(categories[indexes[indexes.length - 1]])
  }
}

/* ===== Formatting Functions ===== */
/*
Format the task names. These are the names of things that begin with '-'. The format
for each of these items is to not be bolded, not have a background, and to have the same
borders as the project area. This function is called in the format_category_names() 
function.

Params: none

Returns: none
*/
function _format_task_names() {
  var indexes = _get_indexes_for_categories();
  var categories = _get_category_names();
  var sheet = _get_project_plan_sheet();
  var gantt_chart = sheet.getRange('gantt_chart');
  for (var i = 0; i < indexes.length; i++) {
    for (var j = indexes[i]; j < indexes[i + 1]; j++) {
      if (j != indexes[i]) {
        gantt_chart.getCell(j + 1, 7).setFontWeight('normal');
        gantt_chart.getCell(j + 1, 7).setFontLine('none');
        var task_range = gantt_chart.offset(j, 6, 1, gantt_chart.getLastColumn() - 7);
        task_range.setBackground('#ffffff');
      }
    }
  }
  for (var i = indexes[indexes.length - 1]; i < gantt_chart.getNumRows() - 1; i++) {
    if (i != indexes[indexes.length - 1]) {
      gantt_chart.getCell(i + 1, 7).setFontWeight('normal');
      gantt_chart.getCell(i + 1, 7).setFontLine('none');
      var task_range = gantt_chart.offset(i, 6, 1, gantt_chart.getLastColumn() - 7);
      task_range.setBackground('#ffffff');
    }
  }
}

/*
Create the inner border of the project area. Called in the format_category_names()
function.

Params: none

Returns: none
*/
function _create_inner_border() {
  var gantt_chart = _get_project_plan_sheet().getRange('gantt_chart');
  var visible_range = gantt_chart.offset(0, 6, gantt_chart.getNumRows(), gantt_chart.getLastColumn() - 7)
  visible_range.setBorder(null, null, null, null, true, true, null, SpreadsheetApp.BorderStyle.DOTTED);
  visible_range.setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  var divider_range = gantt_chart.offset(0, 8, gantt_chart.getNumRows(), 1);
  divider_range.setBorder(null, null, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  divider_range.setBorder(null, true, null, null, null, null, null, null);
  var divider_range = gantt_chart.offset(0, 7, gantt_chart.getNumRows(), 1);
  divider_range.setBorder(null, true, null, null, null, null, null, null);
}

/*
Format the category names, task names, refresh all of the sorting formulas, and protect the
ranges that are not in the project area. This function is called when the 'brush' image is
clicked.

Params: none

Returns: none
*/
function format_category_names() {
  var indexes = _get_indexes_for_categories();
  var gantt_chart = _get_project_plan_sheet().getRange('gantt_chart');
  _create_inner_border();
  for (var i = 0; i < indexes.length; i++) {
    gantt_chart.getCell(indexes[i] + 1, 7).setFontWeight('bold');
    gantt_chart.getCell(indexes[i] + 1, 7).setFontLine('underline');
    var category_range = gantt_chart.offset(indexes[i], 6, 1, gantt_chart.getLastColumn() - 7)
    category_range.setBorder(true, null, null, null, false, null, null, null);
    category_range.setBackground('#f3f3f3');
  }
  _format_task_names();
  _refresh_formulas();
  _protect_ranges();
}

/* ===== Inserting and Deleting Formulas ===== */
/*
Reset the gantt area. This function is used when someone tries to insert a task in
the first row of the project area.

Params: none

Returns: none
*/
function _reset_gantt_chart_range() {
  var gantt_chart = _get_project_plan_sheet().getRange('gantt_chart');
  SpreadsheetApp.getActiveSpreadsheet().setNamedRange('gantt_chart', gantt_chart.offset(-1, 0, gantt_chart.getNumRows() + 1));
}

/*
Insert a row into the project area. The function will not allow a user to insert a
row that is not in the Gantt Chart area.

Params: none

Returns: none
*/
function insert_task() {
  var sheet = _get_project_plan_sheet();
  var cell = sheet.getActiveCell();
  var cell_row = cell.getRow();
  var gantt_chart = sheet.getRange('gantt_chart');
  if (cell_row >= gantt_chart.getLastRow() || cell_row < gantt_chart.getRow() - 1) {
    Browser.msgBox('You can only add rows in the project area.');
  }
  else {
    sheet.insertRowAfter(cell_row);
    if (cell_row == gantt_chart.getRow() - 1) _reset_gantt_chart_range();
    _write_category_names();
    _refresh_formulas();
  }
}

/*
Deletes the row from the project area. The function won't delete the row
if it's not in the project area.

Params: none

Returns: none
*/
function delete_row() {
  var sheet = _get_project_plan_sheet();
  var cell = sheet.getActiveCell();
  var cell_row = cell.getRow();
  var gantt_chart = sheet.getRange('gantt_chart');
  if (cell_row >= gantt_chart.getLastRow() || cell_row < gantt_chart.getRow()) {
    Browser.msgBox('You can only add rows in the project area.');
  }
  else {
    sheet.deleteRow(cell.getRow());
  }
}

/* ===== Sorting Formulas ===== */
/*
Copies all of the category names and sorting formulas down the hidden columns for all of
the categories in the project area.

Params: none

Returns: none
*/
function _refresh_formulas() {
  _write_category_names();
  _copy_sorting_formulas();
  _protect_ranges();
}

/*
Copies the sorting formulas in the hidden columns all the way down the project area. This
function is called in the _refresh_formulas() function.

Params: none

Returns: none
*/
function _copy_sorting_formulas() {
  var sheet = _get_project_plan_sheet();
  var gantt_chart = sheet.getRange('gantt_chart');
  var formula_range = gantt_chart.offset(0, 1, 1, 5);
  formula_range.copyTo(gantt_chart.offset(0, 1, gantt_chart.getNumRows() - 1, 5));
}

/*
Sorts the project area by the hidden sorting columns. This function is called when the
sorting icon is pressed, but it does not have to be called by the user.

Params: none

Returns none
*/
function sort_project_area() {
  var sheet = _get_project_plan_sheet();
  var project = sheet.getRange('gantt_chart');
  _refresh_formulas();
  project.sort([3, 4, 2, 5, 6, 7]);
}

/* ===== Formulas ===== */
/*
Get a dictionary with all of the formulas that you need for sorting the sheet.

Params: 'last_cell' the last cell in the project area {Range}

Returns: 'formulas' a dictionary with all of the formulas {Dictionary}
*/
function _get_formulas(last_cell) {
  var last_cell_a1 = last_cell.getA1Notation();
  var colname = last_cell_a1.substring(0, last_cell_a1.length - 1);
  var row = last_cell_a1.charAt(last_cell_a1.length - 1);
  var formulas = {
    'Category Start': '=minifs($F$7:$F,$B$7:B,B7)',
    'Category End': '=maxifs($G$7:$G,$B$7:B,B7)',
    'Category Integer': '=if(B7=H7,D7*-1,D7)',
    'Task Start': '=iferror(if($K7=3,$K$6,index(K$6:' + colname + '$6,match(1,K7:' + colname + '7,0))),' + colname + '$' + row + ')',
    'Task End': '=iferror(index(K$6:' + colname + '$6,match(3,K7:' + colname + '7,0)),' + colname + '$' + row + ')',
    'Today': '=today()',
    'Days Remaining': '=MAX($K$6:6)-I1'
  }
  return formulas;
}

/* ===== Column Creation, Formatting, and Header Setting ===== */
/*
Create the project area where the chart is going to be. This range will be used
for all project creation purposes.

Params: 'sheet' the sheet being edited {Sheet}, 'start' the start date in integer form,
        {int}, 'end' the end date in integer form {int}
        
Returns: 'last_cell' the final cell in the range of the project {Range}
*/
function _set_gantt_range(sheet, start, end) {
  sheet.getRange('H7').setValue('Phase I');
  sheet.getRange('H8').setValue('- Example task');
  sheet.getRange('H9').setValue('Anchor');
  sheet.getRange('J9').setValue('Only add above this line ^^^^^');
  var last_cell = _create_dates(sheet, start, end);
  var last_column = last_cell.getColumn();
  sheet.getParent().setNamedRange('gantt_chart', sheet.getRange('B7').offset(0, 0, 3, last_column - 1));
  return last_cell;
}

/*
Set the headers for the non-date columns in the Gantt Chart.

Params: 'sheet' the sheet being updated {Sheet}, 'last_cell' the final cell of the range {Range}

Returns: none
*/
function _set_gantt_chart_headers(sheet, last_cell) {
  var header_range = sheet.getRange('B6:J6');
  var header_values = [[
   'Category (hidden)',
   'Category Start',
   'Category End',
   'Category Integer',
   'Task Start',
   'Task End',
   'Task (by Category)',
   'Budget',
   'Description'
  ]];
  header_range.setValues(header_values);
  header_range.setFontWeight('bold');
  header_range.setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  header_range.setBorder(null, null, null, null, true, null);
  /* Set the formulas */
  for (var i = 3; i < 8; i++) {
    sheet.getRange(7, i).setFormula(_get_formulas(last_cell)[header_range.getCell(1, i - 1).getValue()]);
    sheet.getRange(7, i).setNumberFormat('m/d');
  }
  sheet.setColumnWidth(8, 220);
  sheet.setColumnWidth(10, 330);
}

/*
Set the days remaining values cell in the top left of the spreadsheet.

Params: 'sheet' the sheet being edited {Sheet}, 'last_cell' the last cell of the date values
        in the Gantt Chart {Range}
        
Returns: none
*/
function _set_days_remaining_values(sheet, last_cell) {
  var title_range = sheet.getRange('H1:H4');
  title_range.setValues([['Today:'], ['Days Remaining:'], ['In Progress'], ['Complete']]);
  sheet.getRange('I1').setFormula(_get_formulas(last_cell)['Today']);
  sheet.getRange('I2').setFormula(_get_formulas(last_cell)['Days Remaining']);
  sheet.getRange('I3').setBackground('#0000ff');
  sheet.getRange('I4').setBackground('#00ff00');
  var full_range = sheet.getRange('H1:I4');
  full_range.setBorder(false, false, false, false, true, true);
  full_range.setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/*
Create all of the dates at the top of the Gantt chart according to the specified inputs.

Params: 'sheet' the sheet being edited {Sheet}, 'start' the start date of the range as an 
        integer {int}, 'end' the end date of the range as an integer {int}

Returns: none
*/ 
function _create_dates(sheet, start, end) {
  var start_cell = sheet.getRange('K6');
  var cur_date = start;
  var cur_col = 0;
  while (cur_date < end) {
    var cell = start_cell.offset(0, cur_col, 1);
    if (cell.getColumn() == sheet.getMaxColumns()) sheet.insertColumnAfter(cell.getColumn());
    /* Handle border styling */
    if (cur_col == 0) {
      cell.setBorder(true, true, true, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
      cell.setBorder(null, null, null, true, false, false);
    }
    else if (cur_date < end) {
      cell.setBorder(true, false, true, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
      cell.setBorder(null, true, null, true, false, false);
    }
    else if (cur_date + 7 >= end) {
      cell.setBorder(true, false, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
      cell.setBorder(null, true, null, null, false, false);
    }
    cell.setValue(cur_date);
    cell.setNumberFormat('m/d');
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
    sheet.setColumnWidth(cell.getColumn(), 30)
    cur_date += 7;
    cur_col++;
  }
  return start_cell.offset(0, cur_col - 1);
}

/*
Hide all of the columns used for sorting the Gantt Chart area.

Params: 'sheet' the sheet being edited {Sheet}

Returns: none
*/
function _hide_columns(sheet) {
  sheet.hideColumn(sheet.getRange('B:G'));
}

function _format_project_area(sheet) {
  var gantt_chart = sheet.getRange('gantt_chart');
  var format_req = {
    "requests": [{
      "addConditionalFormatRule": { 
        "rule": {
          "ranges": [{
            "sheetId": sheet.getSheetId(),
            "startRowIndex": gantt_chart.getRow() - 1,
            "endRowIndex": gantt_chart.getLastRow() - 1,
            "startColumnIndex": gantt_chart.getColumn() + 2,
            "endColumnIndex": gantt_chart.getLastColumn()
            }],
          "booleanRule": {
            "condition": {
              "type": "TEXT_EQ",
              "values": [{'userEnteredValue': '1'}]
            },
            "format": {
              "backgroundColor": {
                "blue": 1
              },
              "textFormat": {
                'foregroundColor': {
                  "blue": 1
                }
              }
            }
          }
        }, 
        "index": 0
      }
    }, {
      "addConditionalFormatRule": { 
        "rule": {
          "ranges": [{
            "sheetId": sheet.getSheetId(),
            "startRowIndex": gantt_chart.getRow() - 1,
            "endRowIndex": gantt_chart.getLastRow() - 1,
            "startColumnIndex": gantt_chart.getColumn() + 2,
            "endColumnIndex": gantt_chart.getLastColumn()
            }],
          "booleanRule": {
            "condition": {
              "type": "TEXT_EQ",
              "values": [{'userEnteredValue': '3'}]
            },
            "format": {
              "backgroundColor": {
                "green": 1,
              },
              'textFormat': {
                'foregroundColor': {
                  'green': 1
                }
              }
            }
          }
        },
        "index": 1
      }
    }],
  }
  Sheets.Spreadsheets.batchUpdate(JSON.stringify(format_req), sheet.getParent().getId())
}

/*
Freeze the rows and columns of the Gantt Chart area for easy viewing.

Params: 'sheet' the sheet to be edited {Sheet}

Returns: none
*/
function _set_frozen_rows_and_cols(sheet) {
  sheet.setFrozenRows(6);
  sheet.setFrozenColumns(10);
  sheet.setColumnWidth(1, 15);
}
