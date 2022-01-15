#!/home/jaab/PycharmProjects/excel_files_generator/.venv/bin/python

import os
import shutil
import datetime as dt
from factories import mkdir_list, check_file_existance, modify_cells_value

# Defining the variables
actual_year = 2022
root_dir = str(actual_year)
project_path = os.getcwd()
file_to_copy = '/home/jaab/PycharmProjects/excel_files_generator/model.xlsx'

months = [
    'janeiro',
    'fevereiro',
    'marco',
    'abril',
    'maio',
    'junho',
    'julho',
    'agosto',
    'setembro',
    'outubro',
    'novembro',
    'dezembro'
]
week_days = {
    'Mon': 'Segunda',
    'Tue': 'Terca',
    'Wed': 'Quarta',
    'Thu': 'Quinta',
    'Fri': 'Sexta',
    'Sat': 'Sabado',
    'Sun': 'Domingo',
}

# Check if the file to be copied exists
if not check_file_existance(file_to_copy):
    print('Arquivo não encontrado!')
    exit(1)

first_date = dt.date(actual_year, 1, 1)
last_date = dt.date(actual_year, 12, 31)
delta = dt.timedelta(days=1)

# Creating the principal directory and the subdirectories
if root_dir in os.listdir():
    shutil.rmtree(root_dir)

os.mkdir(root_dir)
root_path = os.path.join(project_path, root_dir)
mkdir_list(root_path, months)

# Storing every day of year
all_dates = []
while first_date <= last_date:
    all_dates.append(first_date)
    first_date += delta

# Entering in the root_dir
os.chdir(f'{root_dir}')

# Create the dirs
for index, value in enumerate(months):
    os.chdir(f'{value}/')

    for date in all_dates:
        if date.month == index + 1:
            date_dir = f'{date.strftime("%d-%m-%y")} {week_days[date.strftime("%a")]}'
            os.mkdir(date_dir)

            os.chdir(date_dir)
            file_names = ['Manhã.xlsx', 'Tarde.xlsx']

            for name in file_names:
                shutil.copyfile(os.path.join(project_path, file_to_copy), os.path.join(os.getcwd(), name))
                modify_cells_value(name, 'Sheet1', ['D1', 'J1'], date.strftime("%d/%m/%y"))

            os.chdir('../')

    os.chdir('../')
