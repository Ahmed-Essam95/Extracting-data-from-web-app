# Post Production Test Automation
This project is a Python Selenium-based automation script that dramatically reduces the time required for validating customer billing data in telecom post-production processes.

## Overview
Before automation, the entire process of handling, validating, and extracting customer billing data was done manually. Each dial required checking 22 fields, and with up to 10 dials per account and 400–500 accounts per task, the workload was extremely time-consuming.

After introducing this automated solution using Selenium, the script intelligently interacts with the web-based application to scrape and validate data, reducing manual effort by over 90% and significantly improving accuracy and processing speed.

## Features

- Automated data scraping through web app navigation
- Handles large-scale Excel data writing and processing
- Generates summary reports for quick validation review
- Replaces repetitive manual tasks with a fully automated workflow

## Tools & Technologies

- Python
- Selenium
- openpyxl

## How to Run

-Install the dependencies
-pip install selenium openpyxl
-bash ( python PPT Automation Robot.py )
-Make sure to configure your WebDriver and application login details inside the script before running.

## Author
Ahmed Essam
