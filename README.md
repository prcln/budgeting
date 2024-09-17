# budgeting
This Python-based Budget Management System allows you to automate data management tasks for expense tracking using Google Sheets. It supports functionalities such as retrieving data from Google Sheets, generating PNG images from sheet data, processing and splitting expense data, and managing transactions with undo capabilities.

## Features

- **Retrieve Data**: Fetch data from specified Google Sheets.
- **Generate PNGs**: Create PNG images from Google Sheets data with text wrapping.
- **Process Transactions**: Split and distribute expenses across different sheets based on user-defined rules.
- **Manage Transactions**: Insert new transactions, delete existing rows, and undo recent actions.
- **Clear Sheets**: Clear all data in specified sheets.

## Requirements

- Python 3.x
- `gspread` for Google Sheets API
- `oauth2client` for authentication
- `Pillow` for image processing

## Setup

1. **Clone the Repository**
```
   bash
   git clone https://github.com/yourusername/budget-management-system.git
   cd budget-management-system
```
2. **Install Dependencies**

Ensure you have pip installed, then run:
```
bash
Copy code
pip install gspread oauth2client pillow
```
3. **Google Sheets API Credentials**

Follow this guide to create a Google Sheets API project and obtain credentials.json.
Place the credentials.json file in the project directory.

4. **Update Spreadsheet URL**

Replace the placeholder URL in the update_data() function with your actual Google Sheets URL.

## **Usage**
1. **Run the Script**

Execute the script to start the interactive menu:
```
bash
Copy code
python your_script_name.py
```
2. **Interact with the Menu**

Follow the on-screen prompts to:

- Retrieve data from a sheet
- Generate a PNG from sheet data
- Process and split money data
- Clear all sheets
- List available worksheets
- Insert a new transaction
- Delete a row
- Undo the last action
