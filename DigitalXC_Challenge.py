import openpyxl
import re

excel_path = 'C:\\Users\\Optimus Prime\\PycharmProjects\\AssignmentProject\\coding challenge test.xlsx'
output_file = 'C:\\Users\\Optimus Prime\\PycharmProjects\\AssignmentProject\\coding challenge test output.txt'


# Extracting the groups for "Additional Comments" column
def extract_groups(file_path, string, column_name):

    # Load the Excel file and get the sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Get the header row to find the target column
    headers = [cell.value for cell in sheet[1]]

    if column_name not in headers:
        raise ValueError(f"Column '{column_name}' not found in the Excel file.")

    # Get index of column_name
    column_index = headers.index(column_name)

    # Regex pattern to extract groups
    pattern = re.compile(rf"{string} : \[code\]<I>(.*?)</I>\[/code\]", re.IGNORECASE)
    group_list = []

    for row in sheet.iter_rows(min_row=2, values_only=True):

        # Check if the entire row is empty. Iterates only till the row which has data, not the entire row of sheet
        if all(cell is None for cell in row):
            break
        # Get the cell value from the target column
        cell = row[column_index]
        if cell and isinstance(cell, str) and string in cell:
            match = pattern.search(cell)
            if match:
                # Split groups by commas
                groups = match.group(1).split(",")
                group_list.extend(group.strip() for group in groups if group)

    counts = {}
    for item in group_list:
        if item in counts:
            counts[item] += 1
        else:
            counts[item] = 1
    return counts


def save_to_file(data, file_name):
    with open(file_name, "w") as file:
        file.write(f"{'Group Name':<40}{'Occurrences':>15}\n")
        file.write("=" * 55 + "\n")
        for group, count in data.items():
            file.write(f"{group.title():<40}{count:>15}\n")


def print_the_result(data):
    print(f"{'Group Name':<40}{'Occurrences':>15}")
    print("=" * 55 + "")
    for group, count in data.items():
        print(f"{group.title():<40}{count:>15}")


# Calling the function
input_file = excel_path
output_file = output_file
keyword = "Groups"
column_name = "Additional comments"

try:
    results = extract_groups(input_file, keyword, column_name)
    save_to_file(results, output_file)
    print_the_result(results)
except ValueError as e:
    print(e)
