Skills Development Project 🎓💻
Project Overview
This Python project automates the process of forming student groups and batches based on input data from a Google Sheet. It ensures balanced CGPAs within groups and facilitates better management for educators.

Features
Group Formation:
Automatically creates groups of 4-5 students.
Balances CGPAs among all groups to ensure fairness.
Batch Division:
Divides groups into batches of 3-4 groups for better organization.
Prioritization:
Prioritizes CGPA, followed by other criteria like preferred group members.
File Outputs:
Outputs visually appealing results in Excel and PDF formats.
Technologies Used
Python 🐍
Libraries: pandas, openpyxl, google-api-python-client, PyPDF2
Google Sheets API for fetching input data.
Excel and PDF Formatting for output files.
How It Works
Input:
Data is fetched from a Google Sheet linked to a Google Form filled by students.
Processing:
Students are grouped based on CGPA and other preferences.
Groups are then divided into batches for easier management.
Output:
Results are saved in an Excel file with proper formatting.
A PDF report is generated for distribution.
