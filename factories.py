import os

import openpyxl.utils.exceptions
from openpyxl import load_workbook


def mkdir_list(path, dirs: list):
    for directory in dirs:
        os.makedirs(os.path.join(path, directory))


def check_file_existance(file_path):
    try:
        with open(file_path, 'r'):
            return True
    except FileNotFoundError:
        return False
    except IOError:
        return False


def modify_cells_value(workbook_name: str, sheet_name: str, cells: list, new_value: str):
    try:
        workbook = load_workbook(workbook_name)
        sheet = workbook[sheet_name]

        for cell in cells:
            sheet[cell] = new_value

        workbook.save(workbook_name)

    except openpyxl.utils.exceptions.InvalidFileException:
        return False

    except FileNotFoundError:
        return False
