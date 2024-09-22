# Import necessary libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import random
import textwrap
from PIL import Image, ImageDraw, ImageFont  # For generating the PNG
import time  # For adding delays between requests

# Global variables to store the undo stack and timeout duration
undo_stack = []
undo_timeout_seconds = 60  # Set a 60-second timeout for undo

# Function to retrieve data from Google Sheets
def get_sheet_data(Sheet):
    update_data()
    global spreadsheet
    sheet = spreadsheet.worksheet(Sheet)

    # Get all the data from the sheet
    return sheet.get_all_values()[3:]

# Function to create PNG from Google Sheets data
def generate_image_with_wrapping(sheet_data, output_path):
    # Set image dimensions and colors
    cell_width = 200
    cell_height = 80
    padding = 10
    text_wrap_width = 18
    font_size = 18

    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except IOError:
        font = ImageFont.load_default()

    num_rows = len(sheet_data)
    num_cols = len(sheet_data[0])
    image_width = num_cols * cell_width
    image_height = num_rows * cell_height

    image = Image.new('RGB', (image_width, image_height), 'white')
    draw = ImageDraw.Draw(image)

    for row_idx, row in enumerate(sheet_data):
        for col_idx, cell_value in enumerate(row):
            top_left_x = col_idx * cell_width
            top_left_y = row_idx * cell_height
            bottom_right_x = top_left_x + cell_width
            bottom_right_y = top_left_y + cell_height

            draw.rectangle([top_left_x, top_left_y, bottom_right_x, bottom_right_y], outline='black', width=2)
            wrapped_text = textwrap.fill(str(cell_value), width=text_wrap_width)

            bbox = font.getbbox(wrapped_text)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = top_left_x + padding
            text_y = top_left_y + (cell_height - text_height) // 2

            draw.text((text_x, text_y), wrapped_text, fill='black', font=font)

    image.save(output_path)
    return output_path  # Return the output file path

# Update the data to check if any changes were made
def update_data():
    global transactions_sheet, Person1_sheet, Person2_sheet, Person3_sheet, Person4_sheet, converted_transactions_data, spreadsheet

    # Set up the Google Sheets API credentials and access
    scope = ["https://spreadsheets.google.com/feeds", 
             "https://www.googleapis.com/auth/spreadsheets", 
             "https://www.googleapis.com/auth/drive.file", 
             "https://www.googleapis.com/auth/drive"]

    # Load your credentials.json (Download from Google Console)
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    # Open the spreadsheet
    spreadsheet = client.open_by_url("Your Sheet's URL")

    # Access the sheets, change [Person1] into your worksheet's name
    transactions_sheet = spreadsheet.worksheet("Transactions")
    Person1_sheet = spreadsheet.worksheet("[Person1]")
    Person2_sheet = spreadsheet.worksheet("[Person2]")
    Person3_sheet = spreadsheet.worksheet("[Person3]")
    Person4_sheet = spreadsheet.worksheet("[Person4]")

    # Define expected headers for Transactions sheets
    expected_transactions_headers = ["ID", "Date", "Amount", "Description", "Category", "Note", "Person1", "Person2", "Person3", "Person4", "Number of people", "Split amount"]

    # Get the data from the Transactions sheet
    transactions_data = transactions_sheet.get_all_values()[4:]  # Start reading from row 4

    # Convert the transactions data into dictionaries based on expected headers
    converted_transactions_data = [dict(zip(expected_transactions_headers, row)) for row in transactions_data]

# Process the data and split it across different sheets
def process_splitmoney_data():
    global converted_transactions_data,converted_logic, Person1_sheet, Person2_sheet, Person3_sheet, Person4_sheet

    rows_to_add_Person1 = []
    rows_to_add_Person2 = []
    rows_to_add_Person3 = []
    rows_to_add_Person4 = []

    # Loop through each row in the "Transactions" sheet
    for column in converted_transactions_data:
        ID = column["ID"]
        description = column["Description"]
        split_amount = column['Split amount']
        Person1_involved = column["Person1"].lower() == 'true'  # Checking if it's marked as true
        Person2_involved = column["Person2"].lower() == 'true'
        Person3_involved = column["Person3"].lower() == 'true'
        Person4_involved = column["Person4"].lower() == 'true'
        Note = column["Note"]

        # Initialize default values for transaction_date and category
        transaction_date = column.get("Date")
        category = column.get("Category")

        # Add to each person's sheet based on involvement
        if Person1_involved:
            rows_to_add_Person1.append([ID, transaction_date, split_amount, description, category, Note])
        if Person2_involved:
            rows_to_add_Person2.append([ID, transaction_date, split_amount, description, category, Note])
        if Person3_involved:
            rows_to_add_Person3.append([ID, transaction_date, split_amount, description, category, Note])
        if Person4_involved:
            rows_to_add_Person4.append([ID, transaction_date, split_amount, description, category, Note])

    # Batch update the sheets to reduce API calls
    if rows_to_add_Person1:
        Person1_sheet.append_rows(rows_to_add_Person1)
        time.sleep(2)  # Delay of 2 seconds to avoid API rate limit
    if rows_to_add_Person2:
        Person2_sheet.append_rows(rows_to_add_Person2)
        time.sleep(2)
    if rows_to_add_Person3:
        Person3_sheet.append_rows(rows_to_add_Person3)
        time.sleep(2)
    if rows_to_add_Person4:
        Person4_sheet.append_rows(rows_to_add_Person4)
        time.sleep(2)

# Clear data in sheets to update
def clear_all_sheets():
    global Person1_sheet, Person2_sheet, Person3_sheet, Person4_sheet
    #Modify this if you have more data
    Person1_sheet.batch_clear(["A5:F50"])
    Person2_sheet.batch_clear(["A5:F50"])
    Person3_sheet.batch_clear(["A5:F50"])
    Person4_sheet.batch_clear(["A5:F50"])

#Get a list of available worksheets
def list_worksheets():
    global spreadsheet
    # Fetch all the worksheet objects
    worksheets = spreadsheet.worksheets()

    # Create a list of worksheet titles (names)
    sheet_names = [sheet.title for sheet in worksheets]

    return sheet_names

# Function to generate ID based on the date and row number
def generate_id(date_str, current_row_number):
    # Convert the date string (from column B) to the desired format "YYYYMMDD"
    date_object = datetime.strptime(date_str, '%Y-%m-%d')  # Assuming date format is 'YYYY-MM-DD'
    formatted_date = date_object.strftime('%Y%m%d')  # Convert date to "YYYYMMDD"
    
    # Generate the ID by combining the formatted date, row number, and "E"
    generated_id = f"{formatted_date}-{current_row_number}E"
    return generated_id

# Function to find the first empty row (based on column B being empty)
def find_first_empty_row():
    all_rows = transactions_sheet.get_all_values()[4:]  # Start after the headers

    for i, row in enumerate(all_rows, start=5):  # Start counting from row 5
        # Check if critical fields (Date, Amount, Description) are empty
        if len(row) < 3 or not row[1] or not row[2] or not row[3]:
            return i  # Return the index of the first empty row
    
    # If no empty row is found, return the next available row
    return len(all_rows) + 5

# Function to insert data into the first available empty row
# Modify the function to insert data into the first available empty row
def insert_transactions_row(date, amount, description, category, note, person1_checkbox, person2_checkbox, person3_checkbox, person4_checkbox):
    global undo_stack
    
    # Find the first empty row
    empty_row_number = find_first_empty_row()

    # Generate the ID for column A based on the date (column B) and row number
    id_value = generate_id(date, empty_row_number)

    # Convert checkboxes inputs to 'TRUE'/'FALSE'
    person1_checkbox_value = 'TRUE' if person1_checkbox else 'FALSE'
    person2_checkbox_value = 'TRUE' if person2_checkbox else 'FALSE'
    person3_checkbox_value = 'TRUE' if person3_checkbox else 'FALSE'
    person4_checkbox_value = 'TRUE' if person4_checkbox else 'FALSE'

    # Data to be inserted (ID in column A, then data from column B to J)
    data_to_insert = [
        [id_value,    # ID in column A
        date,        # Date in column B
        amount,      # Amount in column C
        description, # Description in column D
        category,    # Category in column E
        note,        # Note in column F
        person1_checkbox_value, # Person1 checkbox (Column G)
        person2_checkbox_value, # Person2 checkbox (Column H)
        person3_checkbox_value, # Person3 checkbox (Column I)
        person4_checkbox_value]  # Person4 checkbox (Column J)
    ]
    
    # Save current row data for undo purposes (in case of insertion)
    last_row_data = transactions_sheet.row_values(empty_row_number) if transactions_sheet.row_values(empty_row_number) else [''] * 10
    
    # Store the action in the undo stack with the timestamp
    undo_stack.append({
        "action": "insert",
        "row_number": empty_row_number,
        "previous_data": last_row_data,  # In case we want to undo the insert
        "timestamp": time.time()
    })
    
    # Update the first available empty row
    transactions_sheet.update(f'A{empty_row_number}:J{empty_row_number}', data_to_insert, value_input_option='USER_ENTERED')
    print(f"Data inserted successfully into row {empty_row_number} with ID {id_value}.")


# Delete function to store deleted row for undo purposes
def delete_a_row(row_number):
    global undo_stack
    
    # Save the row to be deleted for undo purposes
    last_row_data = transactions_sheet.row_values(row_number)
    
    # Store the action in the undo stack with the timestamp
    undo_stack.append({
        "action": "delete",
        "row_number": row_number,
        "deleted_data": last_row_data,  # In case we want to restore the deleted row
        "timestamp": time.time()
    })
    
    # Delete the row
    transactions_sheet.delete_rows(row_number)
    print(f"Row {row_number} deleted successfully.")

# Function to undo the last action (with timeout and multiple levels)
def undo_last_action():
    global undo_stack
    
    # Get the current time
    current_time = time.time()

    while undo_stack:
        last_action_data = undo_stack.pop()
        action_time = last_action_data["timestamp"]
        
        # Check if the action is within the undo timeout window
        if current_time - action_time <= undo_timeout_seconds:
            row_number = last_action_data["row_number"]

            if last_action_data["action"] == "insert":
                # Undo the last insert by clearing the row
                transactions_sheet.batch_clear([f'A{row_number}:J{row_number}'])
                print(f"Undo insert: Row {row_number} cleared.")
            
            elif last_action_data["action"] == "delete":
                # Undo the last delete by reinserting the deleted data
                deleted_data = [last_action_data["deleted_data"]]
                transactions_sheet.update(f'A{row_number}:J{row_number}', deleted_data, value_input_option='USER_ENTERED')
                print(f"Undo delete: Row {row_number} restored.")
            
            return  # Exit after processing the first valid undo action

        else:
            # If the action is too old to undo, skip it
            print(f"Action at row {last_action_data['row_number']} is too old to undo.")
    
    print("No actions available to undo.")

#Menu for the script
def print_menu():
    print("\nMenu:")
    print("1. Generate PNG from sheet data")
    print("2. Process and split money data")
    print("3. Clear all sheets")
    print("4. List available worksheets")
    print("5. Insert a new transaction")
    print("6. Delete a row")
    print("7. Undo last action")
    print("8. Exit")

def menu():
    while True:
        print_menu()
        choice = input("Enter your choice: ")
        
        if choice == '1':
            sheet_name = input("Enter the sheet name: ")
            output_path = input("Enter the PNG file name (e.g., output_image.png): ")
            sheet_data = get_sheet_data(sheet_name)
            generate_image_with_wrapping(sheet_data, output_path)
            print(f"PNG generated and saved to {output_path}.")
        
        elif choice == '2':
            process_splitmoney_data()
            print("Money data processed and split across sheets.")
        
        elif choice == '3':
            clear_all_sheets()
            print("All sheets cleared.")
        
        elif choice == '4':
            sheet_names = list_worksheets()
            print("Available worksheets:")
            print(sheet_names)
        
        elif choice == '5':
            date = input("Enter the date (YYYY-MM-DD): ")
            amount = input("Enter the amount: ")
            description = input("Enter the description: ")
            category = input("Enter the category: ")
            note = input("Enter the note: ")
            person1_checkbox = input("Is Person1 involved? (yes/no): ").strip().lower() == 'yes'
            person2_checkbox = input("Is Person2 involved? (yes/no): ").strip().lower() == 'yes'
            person3_checkbox = input("Is Person3 involved? (yes/no): ").strip().lower() == 'yes'
            person4_checkbox = input("Is Person4 involved? (yes/no): ").strip().lower() == 'yes'
            insert_transactions_row(date, amount, description, category, note, person1_checkbox, person2_checkbox, person3_checkbox, person4_checkbox)
            print("Transaction inserted successfully.")
        
        elif choice == '6':
            row_number = int(input("Enter the row number to delete: "))
            delete_a_row(row_number)
        
        elif choice == '7':
            undo_last_action()
        
        elif choice == '8':
            print("Exiting...")
            break
        
        else:
            print("Invalid choice. Please enter a number between 1 and 8.")

if __name__ == "__main__":
    # Make sure to update the data before using the menu
    update_data()
    menu()
