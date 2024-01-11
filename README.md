# cilantro-data-transfer-script

## Overview

This Google Apps Script automates the transfer of data between different sheets in different spreadsheets based on predefined settings. The script includes functionalities such as sending data from Source Sheet to Distination Sheet, appending data from Distination Sheet back to Source Sheet, and automating data transfer at specified intervals.

## Usage Instructions

### 1. Send Data from Assignment 2 to Sheet X

#### Function: `sendDataToSheet()`

- This function sends data from the 'Raw Data' sheet in the source spreadsheet to a specified sheet ('Sheet X') in a different spreadsheet.
- Control panel settings in 'Sheet1':
  - Destination file link
  - Sheet name
  - Predefined range

### 2. Append Data from Sheet X back to Source Sheet

#### Function: `AppendDataToSheet()`

- This function appends data from a specified sheet in a different spreadsheet back to the 'Raw Data' sheet in the source spreadsheet (Assignment 2).
- Control panel settings in 'Sheet1':
  - Source file link
  - Sheet name
  - Predefined range

### 3. Automate Data Transfer at Intervals

#### Function: `automateSending()`

- This function allows the automation of data transfer at specified intervals.
- Control panel settings in 'Sheet1':
  - Trigger status (Activated/Disabled)
  - Time interval for automation

## Control Panel Settings

- **Destination File Link (sendDataToSheet):** Specify the URL of the destination spreadsheet.
- **Sheet Name (sendDataToSheet & AppendDataToSheet):** Specify the name of the sheet to send/append data.
- **Predefined Range (sendDataToSheet & AppendDataToSheet):** Specify the predefined range of data.
- **Source File Link (AppendDataToSheet):** Specify the URL of the source spreadsheet.
- **Time Interval for Automation (automateSending):** Specify the time interval in hours for automated data transfer.

## Notes

- The code includes functions to preserve font color, font family, and font weight settings during data transfer.
- Control panel settings should be configured before running the respective functions.
- The automation trigger can be activated or deactivated using the `automateSending` function.

## Important

- Ensure that the required permissions are granted for the script to access and modify spreadsheets.

