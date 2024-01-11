# cilantro-data-transfer-script

## Overview

This Google Apps Script automates the transfer of data between different sheets in different spreadsheets based on predefined settings. The script includes functionalities such as sending data from Assignment 2 to Sheet X, appending data from Sheet X back to Assignment 2, and automating data transfer at specified intervals.

## Usage Instructions

### 1. Send Data from Assignment 2 to Sheet X

#### Function: `sendDataToSheet()`

- This function sends data from the 'Raw Data' sheet in the source spreadsheet (Assignment 2) to a specified sheet ('Sheet X') in a different spreadsheet.
- Control panel settings in 'Sheet1':
  - Destination file link: Cell B1
  - Sheet name: Cell B2
  - Predefined range: Cell B3

### 2. Append Data from Sheet X back to Assignment 2

#### Function: `AppendDataToSheet()`

- This function appends data from a specified sheet ('Sheet X') in a different spreadsheet back to the 'Raw Data' sheet in the source spreadsheet (Assignment 2).
- Control panel settings in 'Sheet1':
  - Source file link: Cell B5
  - Sheet name: Cell B6
  - Predefined range: Cell B7

### 3. Automate Data Transfer at Intervals

#### Function: `automateSending()`

- This function allows the automation of data transfer at specified intervals.
- Control panel settings in 'Sheet1':
  - Trigger status (Activated/Disabled): Cell B10
  - Time interval for automation: Cell B11

## Control Panel Settings

- **Destination File Link (sendDataToSheet):** Specify the URL of the destination spreadsheet in Cell B1.
- **Sheet Name (sendDataToSheet & AppendDataToSheet):** Specify the name of the sheet to send/append data in Cells B2/B6.
- **Predefined Range (sendDataToSheet & AppendDataToSheet):** Specify the predefined range of data in Cells B3/B7.
- **Source File Link (AppendDataToSheet):** Specify the URL of the source spreadsheet in Cell B5.
- **Time Interval for Automation (automateSending):** Specify the time interval in hours for automated data transfer in Cell B11.

## Notes

- The code includes functions to preserve font color, font family, and font weight settings during data transfer.
- Control panel settings should be configured before running the respective functions.
- The automation trigger can be activated or deactivated using the `automateSending` function.

## Important

- Ensure that the required permissions are granted for the script to access and modify spreadsheets.

