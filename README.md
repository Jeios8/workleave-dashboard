# Employee Work Leave Tracker

## Overview
This project is an Excel-based application designed to help track the work leave and vacation schedules of employees, including holidays. It dynamically updates employee leave records, provides team and individual employee dashboards, and ensures accurate tracking of holidays.

The project includes a **VBA (Visual Basic for Applications)** script that automates the update of data in various tables, streamlining the process and improving efficiency.

## Features
- **Dynamic Updates:** The VBA script automatically updates leave records and dashboards when changes are made.
- **Employee Dashboard:** View and track individual employee leave records and vacation days.
- **Team Dashboard:** Get a bird’s eye view of team leave schedules to avoid overlap and ensure smooth operations.
- **Holiday Tracking:** Automatically accounts for public and custom holidays, adjusting leave calculations accordingly.

## Requirements
- **Microsoft Excel** (2010 or later)
- **Enable Macros:** Macros must be enabled for the VBA scripts to function correctly.

## How to Use
1. **Open the Excel file** and make sure to enable macros when prompted.
2. **Employee Report:** Navigate to the Employee Dashboard to add or edit employee information and leave requests.
3. **Team Dashboard:** Use the Team Dashboard to view team-wide leave schedules and identify any overlapping leave periods.
4. **Holidays:** You can add or edit public holidays in the `Holiday Table`, and the system will adjust employee leave calculations accordingly.

## VBA Script Functions
The VBA script included in this project performs the following:
- **Automatic Table Updates:** Ensures that any new entries in employee or leave data are reflected in the dashboards.
- **Leave Calculations:** Dynamically recalculates the number of vacation days based on the holidays and current leave entries.
- **Error Handling:** Provides basic error handling to alert users when input data is missing or incorrect.

## Getting Started
1. Download the Excel file and open it in Excel.
2. Enable macros when prompted.
3. Add employee data and leave requests using the dashboards.
4. Review team-wide leave schedules and make adjustments as needed.

## Customization
The application is designed to be flexible and can be customized to fit your organization's needs:
- **Additional Leave Types:** You can add new types of leaves (e.g., sick leave, maternity leave) by modifying the relevant tables.
- **Team Structures:** Easily adjust the Team Dashboard to reflect organizational changes.
- **Holiday Adjustments:** Modify the holiday table to accommodate different regional or company-specific holidays.

## License
This project is licensed under the MIT License. See the `LICENSE` file for more details.

## Author
Developed by Jeios8
