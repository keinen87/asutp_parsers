import os
import json
import pandas
import re
import time
from pprint import pprint
from common import create_doc_final

LOGS_PATH = 'logs'
MONTHS = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь',
]
TEMPLATE_PATH = 'шаблон.docx'


def get_files(path):
    finded_filenames = []

    logs_dir_filenames = os.listdir(path)
    for month in MONTHS:
        for filename in logs_dir_filenames:
            if month in filename:
                finded_filenames.append((filename, month))
    return finded_filenames


if __name__ == '__main__':
    report = []
    year_weight_sum = 0
    disp_logs = get_files(LOGS_PATH)
    for disp_log, month in disp_logs:
        print(disp_log)
        month_weight_sum = 0
        days = []
        for page in range(1, 32):
            try:
                if page in range(1, 10):
                    page = f'0{page}'
                log = pandas.read_excel(
                    os.path.join(LOGS_PATH, disp_log),
                    usecols='W:AG',
                    header=49,
                    nrows=1,
                    sheet_name=f'{page}',
                    na_filter=False).to_dict(orient='records'
                    )
                if page not in (10, 20, 30):
                    page = str(page).replace('0', '')
                log = log[0]
                match = False
                for record in log:
                    try:
                        if 'Римпул' in record:
                            match = True
                            target = re.findall(r'\d*,\d*', record)
                            if not target:
                                target = re.findall(r'\d{3,}', record)
                                print(target)  
                            try:
                                weight = float(target[0].replace(',', '.'))
                                month_weight_sum += weight
                                year_weight_sum += weight
                                print(f'день {page} -- {weight}')
                                day = {
                                    page: weight
                                }
                                days.append(day)
                            except IndexError:
                                weight = 0
                                print(f'день {page} -- {weight}')
                                day = {
                                    page: weight
                                }
                                days.append(day)
                                continue  
                    except TypeError:
                        continue
                if not match:
                    weight = 0
                    print(f'день {page} -- {weight}')
                    day = {
                        page: weight
                    }
                    days.append(day)    
            except ValueError:
                continue
        month_report = {
            month: {
                'days': days,
                'month_weight_sum': round(month_weight_sum, 3)
            }
        }
        report.append(month_report)    
    print(year_weight_sum)
    context = {
        'report': report,
        'year_weight_sum': round(year_weight_sum, 3)
    }
    try:
        create_doc_final(TEMPLATE_PATH, context, 'disp_report')
    except PermissionError:
        create_doc_final(TEMPLATE_PATH, context, 'disp_report', exists=True)

    with open('disp_report.json', 'w', encoding='utf8') as file:
        json.dump(context, file, ensure_ascii=False)
