# Simple Gantt Chart for Google Sheets

Gantt Chart's shouldn't be hard to put together quickly, but can take a lot of time to do on your own in an Excel or Google Sheet. With a few steps, this Google Script lets you build a Gantt chart without having to sign up for new software, download anything, or pay for anything.

# Requirements

* A Google Account

# Steps to Install

1. Open the script editor on your Google Sheet:
2. Add the advanced Google Sheet API in the Script Editor:
3. Add the Google Sheet API in your API console:
4. Paste the code from `GanttChart_v2.gs` in this repository into your script editor (doesn't matter what you call the file).
5. Run the `create_sheet` function in the script
6. Insert all of the images from this repository into your Sheet.
7. Assign the following scripts to each image (can copy/paste the function names):
  * Paintbrush - `format_category_names`
  * Add - `insert_task`
  * Sort - `sort_project_area`
  * Trash - `delete_row`
  
# Instructions for Use
