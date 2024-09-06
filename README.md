
# Timecard Automation Script

This script automates timecard entries by fetching dates from an Excel file and submitting them to Tickspot. The script supports filtering dates by a specified range and works with multiple sheets in the Excel file.

## Prerequisites

- Python 3.x
- An account on Tickspot
- Excel file with timecard entries in the following format:
  - Each sheet represents a person.
  - Each sheet has a `date` column (first column), with dates in `dd/mm/yyyy` format.

## Setup

### Step 1: Clone the repository

`git clone <repository-url>`
`cd <repository-directory>`

### Step 2: Create a virtual environment

Use the following commands to create a Python virtual environment:

`python3 -m venv venv`

### Step 3: Activate the virtual environment

On **macOS/Linux**:

`source venv/bin/activate`

On **Windows**:

`venv\Scripts\activate`

### Step 4: Install dependencies

Make sure you have a `requirements.txt` file with the necessary dependencies. Use the following command to install them:

`pip install -r requirements.txt`

Your `requirements.txt` should contain the following dependencies:

`requests`
`beautifulsoup4`
`pandas`
`openpyxl`

## Usage

The script requires several command-line arguments to run:

- **--email**: Your Tickspot email for login.
- **--password**: Your Tickspot password for login.
- **--file**: Path to the Excel file containing the dates.
- **--sheet**: The sheet name representing the person in the Excel file.
- **--start_date**: The start date to filter the entries (format: `dd/mm/yyyy`).
- **--end_date**: The end date to filter the entries (format: `dd/mm/yyyy`).
- **--project**: The project name (optional, defaults to `MTK`).
- **--hours**: Number of hours to load per day (optional, defaults to 8).

### Example

`python main.py --email your_email@example.com --password your_password --file ./path_to_excel_file.xlsx --sheet JohnDoe --start_date 01/09/2024 --end_date 30/09/2024 --project MTK --hours 8`

### What the script does:

1. Logs into your Tickspot account using the provided credentials.
2. Fetches the dates from the specified Excel sheet and filters them based on the `start_date` and `end_date` range.
3. Submits timecard entries for the filtered dates with the specified number of hours to the given project.

### Confirmation

Before submitting time entries, the script will list the dates that will be processed and ask for your confirmation. Enter `yes` to proceed, or `no` to cancel.

## Deactivation

When you are done using the virtual environment, deactivate it by running:

`deactivate`