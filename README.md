Automated Audit and Email Automation

This repository contains VBA macros designed to automate auditing tasks and email generation using Microsoft Excel and Outlook. The primary function of these macros is to perform audits on specified criteria within an Excel spreadsheet and automatically generate an email with the audit results for review by stakeholders.

Overview

When the Send_All macro is executed, it prompts the user to generate an audit. If the user confirms, the Ost_Loop macro is called, which loops through the specified criteria and generates an email with the audit findings.

Prerequisites

To use these macros, ensure that you have the following:

- Microsoft Excel installed
- A local copy of Microsoft Outlook installed

How It Works

1. Send_All Macro:

- Prompts the user to generate an audit.
- If confirmed, calls the Ost_Loop macro.
Upon completion, displays a message indicating the audit is complete.

2. Ost_Loop Macro:

- Filters data in the "AUDIT" worksheet based on specified criteria.
- Copies filtered data to a new worksheet named "PlaceHolder".
- Creates an Outlook email with the audit findings and displays it for review.
- Deletes the temporary worksheet and resets the filter.

Usage

1. Open the Excel workbook containing the data to be audited.
2. Enable macros if prompted.
3. Execute the Send_All macro.
4. Confirm whether to generate the audit.
5. Review the generated email in Microsoft Outlook. 

Files

- Send_All.bas: Contains the Send_All macro responsible for initiating the audit process.
- Ost.bas: Contains the Ost_Loop macro, which performs the audit and generates the email.
- RangetoHTML Function: Utility function to convert a range to HTML for email formatting.

Note

Ensure that the email recipient and subject are appropriately configured within the Ost_Loop macro before running the audit.