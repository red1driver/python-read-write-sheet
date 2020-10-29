# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import configparser
import logging
import os


_dir = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config.read(_dir + r"\config\config.ini")
api_token=config['api_token']['token']
#print(api_token)


# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
column_map = {}


# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)


# TODO: Replace the body of this function with your code
# This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
#
# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_row_and_build_updates(source_row):
    # Find the cell and value we want to evaluate
    status_cell = get_cell_by_column_name(source_row, "STREET ADDRESS/ SEGMENT ID")
    status_value = status_cell.display_value
    if status_value == "3164 BONA ST":
        remaining_cell = get_cell_by_column_name(source_row, "CUSTOMER CONTACT INFO")
        if remaining_cell.display_value != "0":  # Skip if already 0
            print("Need to update row #" + str(source_row.row_number))

            # Build new cell value
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map["CUSTOMER CONTACT INFO"]
            new_cell.value = "Yes"
            new_cell.strict = False

            # Build the row to update
            new_row = smart.models.Row()
            new_row.id = source_row.id
            new_row.cells.append(new_cell)

            #print(new_row)
            return new_row

    return None



"""
#print("Loaded " + str(len(input_sheet.rows)) + " rows from sheet: " + input_sheet.name)
#folders = smart.Home.list_folders(include_all=True)
#print(folders)


print("Starting ...")

# Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
smart = smartsheet.Smartsheet(access_token=api_token)
# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
#logging.basicConfig(filename='rwsheet.log', level=logging.INFO)
logging.basicConfig(filename=_dir + '/rwsheet.log', level=logging.INFO)
# Import the sheet
result = smart.Sheets.import_xlsx_sheet(_dir + '/Sample input_sheet.xlsx', header_row_index=0)

# Load entire sheet
sheet = smart.Sheets.get_sheet(result.data.id)

"""


print("Starting ...")

# Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
smart = smartsheet.Smartsheet(access_token=api_token)
# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
#logging.basicConfig(filename='rwsheet.log', level=logging.INFO)
logging.basicConfig(filename=_dir + '/rwsheet.log', level=logging.INFO)
# Import the sheet


# OPEN YOUR SHEETS

target_folder_id = 2776500131915652   # Sheets/Daniel/Dev
target_sheet_id = 8951278194714500  # Sheets/Daniel/Dev/CA_DEV

# import xls file into new smartsheet
input_result = smart.Folders.import_xlsx_sheet(
    target_folder_id,  # folder id
    _dir + '/CA_dev.xlsx',
    header_row_index=0
)

# imported sheet object
input_sheet = smart.Sheets.get_sheet(input_result.data.id)

# update target sheet object
update_target = smart.Sheets.get_sheet(target_sheet_id)



print("Loaded " + str(len(input_sheet.rows)) + " rows from sheet: " + input_sheet.name)



# Build column map for later reference - translates column names to column id
for column in input_sheet.columns:
    column_map[column.title] = column.id

# Accumulate rows needing update here
rowsToUpdate = []

for row in input_sheet.rows:
    rowToUpdate = evaluate_row_and_build_updates(row)
    if rowToUpdate is not None:
        rowsToUpdate.append(rowToUpdate)

# Finally, write updated cells back to Smartsheet
if rowsToUpdate:
    print("Writing " + str(len(rowsToUpdate)) + " rows back to sheet id " + str(input_sheet.id))
    input_result = smart.Sheets.update_rows(target_sheet_id, rowsToUpdate)
    #input_result = smart.Sheets.update_rows(input_result.data.id, rowsToUpdate)
else:
    print("No updates required")

print("Done")
