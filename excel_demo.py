import openpyxl

print('All inputs are case sensitive!')
file_name = str(input("Enter a file name: "))
file_type = '.xlsx'
concat = file_name + file_type

wb = openpyxl.load_workbook(concat)
ws = wb.active

print('Command list: modify, review, titles, get, save.\n'
      'Remember to save after finishing with your changes.')

user_input = input("Command: ")

command_list = 'modify', 'review', 'titles', 'save', 'get',
sheet_list = wb.sheetnames


def get_sheet():
    sheet_input = input('Enter sheet name: ')
    print(wb[sheet_input])


def review_sheets():
    """Prints all sheet names"""
    print(wb.sheetnames)


def staff_contact():
    """Enter user input into cells."""
    phone = 10
    data = input('Enter phone number: ')
    if phone == data:
        ws['D4'] = data
    elif len(data) != phone:
        display_error()


def save_doc():
    """Saves document changes."""
    wb.save(concat)


def worksheet_titles():
    """Loop through worksheets and rename each one."""
    for sheet in wb:
        first_cell_value = str(sheet['A2'].value)
        sheet.title = first_cell_value
        print(sheet.title)


def display_error():
    print('Make a valid entry.')


# While loop takes user input and executes functions.
while user_input.lower() != 'q':
    if user_input == 'modify':
        staff_contact()
    elif user_input == 'review':
        review_sheets()
    elif user_input == 'titles':
        worksheet_titles()
    elif user_input == 'save':
        save_doc()
    elif user_input == 'get':
        get_sheet()
    elif user_input != command_list:
        print('Please enter a valid command.')

    user_input = input("Command: ")
print('Good bye!')