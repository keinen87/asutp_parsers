import csv
import json
import os
import time
from datetime import datetime
from collections import defaultdict
from pprint import pprint
from terminaltables import SingleTable
from common import create_doc_final

FOLDER_PATH = r'\\192.168.25.97\c\Logs\Scale\Trunk_conveyor' # r'logs'
# FILENAME = 'Trunk_scale_weight_hour_log_0.csv'
FILENAME = 'Trunk_scale_weight_minute_log_0.csv'
TEMPLATE_PATH = 'шаблон.docx'


YEAR = 2024
MONTH = 1
DAY = 1
EVENTS_QTY = 15


def get_events_table(report):
    table_data = []
    table_data.append(
        [YEAR, f"{report['year_weight_sum']} т."]
    )
    for month_report in report['report']:
        for month_detail in month_report:
            table_data.append(
                [month_detail, f"{month_report[month_detail]['month_weight_sum']} т."]
            )
            # Если нужно по дням
            # for day in month_report[month_detail]['days']:
            #     day_num = list(day.items())[0][0]
            #     day_weight = list(day.items())[0][1]
            #     table_data.append(
            #         [day_num, day_weight]
            #     )
    table_instance = SingleTable(table_data)
    table_instance.inner_heading_row_border = True
    return table_instance.table


def get_parsed_log(path):
    parsed_log = []
    with open(path, 'r', newline='') as file:
        reader = csv.DictReader(file, delimiter=';')
        for row in reader:
            log_row_datetime = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S')
            log_row_date = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S').date()
            log_row_weight = round(float(row['VarValue'].replace(',', '.')), 3)
            if log_row_date.year == YEAR:
                parsed_log.append({
                    'log_row_date': log_row_date,
                    'log_row_time': log_row_datetime,
                    'log_row_weight': log_row_weight
                })
    return parsed_log  


def get_full_report(parsed_log):
    events = defaultdict(list)
    for event in parsed_log:
        events[event['log_row_date']].append(event)

    for event in events:
        day_weight_sum = 0
        for event_info in events[event]:
            day_weight_sum += event_info['log_row_weight']  
        events[event].append({'day_weight_sum': round(day_weight_sum/1000, 3)})

    months = {
        1: 'Январь',
        2: 'Февраль',
        3: 'Март',
        4: 'Апрель',
        5: 'Май',
        6: 'Июнь',
        7: 'Июль',
        8: 'Август',
        9: 'Сентябрь',
        10: 'Октябрь',
        11: 'Ноябрь',
        12: 'Декабрь',
    }
    report = []
    year_weight_sum = 0
    for month in months:
        days = []
        month_weight_sum = 0
        for date_event in events:
            if month == date_event.month:
                day_weight_sum = events[date_event][-1]['day_weight_sum']
                day = {
                    date_event.day: day_weight_sum
                }
                days.append(day)
                month_weight_sum += day_weight_sum
            else:
                continue
        month_weight_sum = round(month_weight_sum, 3)
        year_weight_sum += month_weight_sum
        month_report = {
            months[month]: {
                'days': days,
                'month_weight_sum': month_weight_sum
            }
        }
        report.append(month_report)
    full_report = {
        'report': report,
        'year_weight_sum': round(year_weight_sum, 3),
        'current_year': YEAR
    }
    return full_report


def main():
    filepath = os.path.join(FOLDER_PATH, FILENAME)
    parsed_log = get_parsed_log(filepath)
    full_report = get_full_report(parsed_log)
    print(get_events_table(full_report))
    try:
        create_doc_final(TEMPLATE_PATH, full_report, 'scales_report')
    except PermissionError:
        create_doc_final(TEMPLATE_PATH, full_report, 'scales_report', exists=True)
    with open('scales_report.json', 'w', encoding='utf8') as file:
        json.dump(full_report, file, ensure_ascii=False)


if __name__ == '__main__':
    main()
