#  Player Groupings and CSV Export Automation

[See Code here](https://github.com/harrybehn/PortfolioProjectCode/blob/main/Program%20Email%20sending.sas)

## Overview
This project automates the process of grouping players by Food Credits and generating separate CSV files for each group using VBA in Excel. It is designed to streamline manual data segmentation and export tasks, making it ideal for operational reporting or incentive tracking.

## Features
- Reads player data from a specified worksheet.
- Groups players by their Food Credits using a dictionary structure.
- Creates a new workbook for each group and populates it with relevant PlayerIDs.
- Saves each workbook as a CSV file named after the Food Credit tier.
- Includes error handling for missing sheets or unexpected issues.

## Process Flow
1. Read Inputs
  - Sheet name and file name are read from specific cells in the control sheet (Sheet1).
2. Group Players
  - A dictionary is used to map each Food Credit value to a collection of PlayerIDs.
3. Generate CSVs
  - For each Food Credit group:
   - A new workbook is created.
   - PlayerIDs are written to the sheet.
   - The file is saved as Food Credits - [Tier].csv.
4. Error Handling
  - Alerts the user if the sheet name is incorrect or if any unexpected error occurs.
