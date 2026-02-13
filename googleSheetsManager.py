import datetime
from datetime import date
from gspread_formatting import *
from config import *

def main():
    add_new_recruits()

def clean_data(data_list):
    for index, row in enumerate(data_list[:]):
        if len(row) < len(data_list[0]):
            data_list.remove(row)

def get_phone_numbers() -> list:
    global sheet_id, scopes, creds, client

    phone_numbers = []

    workbook = client.open_by_key(sheet_id)
    worksheet = workbook.worksheets()[1]

    last_row = len(worksheet.col_values(1))
    data = worksheet.get(f"A4:M{last_row}")
    clean_data(data)

    absent_phone_number_message = "Не надав номер телефону"

    for row in data:
        if row[-1] != "Тимчасово неактивний":
            name = row[2] if row[2] != "" else row[1]
            phone_number = row[4] if row[4] != "-" else absent_phone_number_message
            phone_numbers.append([name, phone_number])
    return phone_numbers

def format_phone_numbers_to_print(phone_numbers: list[str]) -> list[str]:
    ph = phone_numbers[:]
    for i, num in enumerate(phone_numbers):
        s = num[1]
        s = s.replace("(", "").replace(")", "")
        s = s.replace(" ", "")
        s = "+38" + s
        ph[i][1] = s
    return ph

def format_phone_number(phone_number: str):
    if phone_number[0] == "+":
        phone_number = phone_number.removeprefix("+38")
    elif phone_number[0] == "3":
        phone_number = phone_number.removeprefix("38")
    elif phone_number[0] != "0":
        phone_number = "0" + phone_number
    elif phone_number[0] == "(":
        return phone_number

    phone_number = ("(" + phone_number[:3] + ") "
                    + phone_number[3:6] + " " + phone_number[6:8] + " "
                    + phone_number[8:])
    return phone_number

def format_blood_group(blood_group: str) -> str:
    if blood_group[0].isdigit():
        match blood_group[0]:
            case '1':
                blood_group = blood_group.removeprefix('1')
                blood_group = 'O (I) '+ blood_group
            case '2':
                blood_group = blood_group.removeprefix('2')
                blood_group = 'A (II) ' + blood_group
            case '3':
                blood_group = blood_group.removeprefix('3')
                blood_group = 'B (III) ' + blood_group
            case '4':
                blood_group = blood_group.removeprefix('4')
                blood_group = 'AB (IV) ' + blood_group
    elif blood_group[0] == 'I':
        s = "("
        for i in blood_group:
            if not i.isalpha():
                s += ')' + i
        return s
    return blood_group

def fit_data_to_members_sheet_format(data: list[str]) -> list[str, int]:
    global sheet_id, scopes, creds, client
    workbook = client.open_by_key(sheet_id)
    members_sheet = workbook.worksheets()[1]

    (name, callsign, birth_date,
     phone_number, address, blood_group,
     education, telegram_tag, acknowledgement, date_of_joining) = (
        data[1], data[2], data[3],
        data[4], data[5], data[6],
        data[7], data[8], data[9], data[0]
    )

    phone_number = format_phone_number(phone_number)
    date_of_joining = date_of_joining.split(' ')[0]

    blood_group = format_blood_group(blood_group)

    current_year = date.today().year
    members_birth_dates = members_sheet.col_values(4)

    last_indexes = [0 for i in range(3)]

    for index, d in enumerate(members_birth_dates):
        try:
            if not datetime.datetime.strptime(d, "%d.%m.%Y"):
                continue
        except:
            pass
        else:
            if current_year-int(d.split(".")[-1]) in range(18, 26):
                last_indexes[0] = index
            elif current_year-int(d.split(".")[-1]) in range(16, 18):
                last_indexes[1] = index
            elif current_year-int(d.split(".")[-1]) in range(14, 16):
                last_indexes[2] = index

    row_index = 100
    recruit_birth_date_year = int(birth_date.split(".")[-1])
    if current_year - recruit_birth_date_year in range(18, 26):
        row_index = last_indexes[0]
    elif current_year - int(d.split(".")[-1]) in range(16, 18):
        row_index = last_indexes[1]
    elif current_year - int(d.split(".")[-1]) in range(14, 16):
        row_index = last_indexes[2]
    row_index += 2  # +2 because in google sheets indexation starts with 1 and +1 to set to the next row

    num = int(members_sheet.col_values(1)[row_index-2])+1

    fitted_data = [str(num), name, callsign,
                   birth_date, phone_number, address,
                   blood_group, '', "Прихильник", education,
                   acknowledgement, date_of_joining, '']
    return fitted_data, row_index

def add_new_recruits():
    global sheet_id, scopes, creds, client
    workbook = client.open_by_key(sheet_id)
    recruit_sheet = workbook.worksheets()[2]

    row_to_add = len(recruit_sheet.col_values(1))
    recruit_data_to_add = recruit_sheet.row_values(row_to_add)

    members_sheet = workbook.worksheets()[1]
    fitted_data, row_index = fit_data_to_members_sheet_format(recruit_data_to_add)

    members_sheet.insert_row(fitted_data, row_index)

    style_cells(members_sheet, row_index)

def style_cells(members_sheet, row):
    # Cell styles
    number_cell_color = Color(217 / 255, 234 / 255, 211 / 255)
    number_cell = CellFormat(
        backgroundColor=number_cell_color,
        horizontalAlignment='CENTER'
    )

    border = Border(style='SOLID', width=1, color=Color(0, 0, 0))

    row_format = CellFormat(
        borders=Borders(
            top=border,
            bottom=border,
            left=border,
            right=border
        ),
        verticalAlignment='MIDDLE',
        textFormat=TextFormat(bold=False)
    )

    center_format = CellFormat(
        horizontalAlignment='CENTER',
    )

    last_row = row
    format_cell_range(members_sheet, f"A{last_row}", number_cell) # Number cell
    format_cell_range(members_sheet, f"A{last_row}:M{last_row}", row_format) # Entire row
    format_cell_range(members_sheet, f"C{last_row}:E{last_row}", center_format) # Center cells
    format_cell_range(members_sheet, f"G{last_row}:I{last_row}", center_format) # Again
    format_cell_range(members_sheet, f"K{last_row}:L{last_row}", center_format) # Again

    # Chip cell style
    request = {
        "copyPaste": {
            "source": {
                "sheetId": members_sheet.id,
                # Coords of I4
                "startRowIndex": 3, "endRowIndex": 4,
                "startColumnIndex": 8, "endColumnIndex": 9
            },
            "destination": {
                "sheetId": members_sheet.id,
                "startRowIndex": last_row - 1, "endRowIndex": last_row,
                "startColumnIndex": 8, "endColumnIndex": 9
            },
            "pasteType": "PASTE_NORMAL",
            "pasteOrientation": "NORMAL"
        }
    }

    members_sheet.spreadsheet.batch_update({"requests": [request]})

    default_choice = "Прихильник"
    members_sheet.update_acell(f"I{last_row}", default_choice)
