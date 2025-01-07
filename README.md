# Selenium-Automation_Data_Scraping

This repository contains a Python script for automating data scraping and processing tasks using Selenium, Pandas, and other libraries. The script handles downloading, processing, and organizing data files from a secure web portal.

## Features

1. **Web Automation**:
   - Automates login to a secure web portal.
   - Downloads reports or data based on predefined date ranges.
   - Dynamically selects locations and types of reports.

2. **File Management**:
   - Creates a structured folder hierarchy (`YYYYMM/DD`) for storing downloaded files.
   - Cleans up old files in the specified directory.
   - Renames files with meaningful names (e.g., location and report type).

3. **Data Processing**:
   - Processes and validates CSV files using Pandas.
   - Handles blank rows and standardizes column names.
   - Ensures data is clean and ready for further use.

4. **Email Notifications**:
   - Sends email alerts for successes or failures.
   - Covers scenarios like login failure, empty files, or processing errors.

5. **Logging**:
   - Tracks all operations and errors in a custom log for easy troubleshooting.

## Prerequisites

- Python 3.8 or higher
- Google Chrome (latest version)
- ChromeDriver (managed automatically with `webdriver_manager`)
- Required Python libraries (install via `pip install -r requirements.txt`)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/DebasisNandiCode/Selenium-Automation_Data_Scraping.git
