#  Player Groupings and CSV Export Automation

[See VBA Code here](https://github.com/harrybehn/PortfolioProjectCode/blob/main/PlayerGroupings.bas)
[See Excel File here](https://hannresortsphil-my.sharepoint.com/:x:/g/personal/harry_francisco_hannresorts_com/EVUwemcHuABMiBCyBd41q28Bv59fOhqD5dhslqHuieu2bQ?e=2d9ZB2)

## Overview
This project automates the process of grouping players by Food Credits and generating separate CSV files for each group using VBA in Excel. It is designed to streamline manual data segmentation and export tasks, making it ideal for operational reporting or incentive tracking.

## Features
- Reads player data from a specified worksheet.
- Groups players by their Food Credits using a dictionary structure.
- Creates a new workbook for each group and populates it with relevant PlayerIDs.
- Saves each workbook as a CSV file named after the Food Credit tier.
- Includes error handling for missing sheets or unexpected issues.

## Process Flow

1. **Read Inputs**
   - Retrieve the sheet name and base file name from specific cells in the control sheet (`Sheet1`).

2. **Group Players**
   - Use a dictionary to map each unique Food Credits value to a collection of PlayerIDs.
   - Loop through the raw data and populate the dictionary accordingly.

3. **Generate CSV Files**
   - For each Food Credits group:
     - Create a new workbook.
     - Write the PlayerIDs to the first column.
     - Save the workbook as a CSV file named `Food Credits - [Tier].csv`.

4. **Error Handling**
   - Display a message if the sheet name is incorrect or if any unexpected error occurs.

## Technologies Used
- MS Excel
- VBA (Visual Basic for Applications)

