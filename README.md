# MasterEmailManager
This script and its accompanying spreadsheet (at <https://docs.google.com/spreadsheets/d/1adBAqaA5qVB2awvPCf9rc-O6a1S8fzZetUuM6uE5Osc>) require you to have a Google account before you can use them.  Furthermore, this script requires you to enable Gmail Advanced Service in your Google Apps Script.  Do not request editing permissions for the primary script or the deployment spreadsheet.  Instead, make a copy of the script and spreadsheet to your Google Apps Script environment and Google Drive, respectively.  

This Google Script automates email archiving, retention, and deletion.  It archives emails after a user-specified number of hours and permanently deletes emails after a user-specified number of days.  The accompanying spreadsheet is an example of a spreadsheet with Gmail search parameters the script can use.  If you make a copy of the example spreadsheet, you can list Gmail search parameters of your choice in each cell in Column A.

There are time triggers that you need to set up first.  The script comes with the source code for time triggers that you can run in Google Apps Script.  
1. Run the `createTimeTriggerEveryNHours` function to set up the automation schedule.
2. Run the `MasterEmailManager` function once to start regular automation.

If you must modify your code, follow the below instructions.
1. Run the `stopTriggers` function.
2. Make your edits and save the script.
3. Run the `createTimeTriggerEveryNHours` function.
4. Run the `MasterEmailManager` function to resume automation.
