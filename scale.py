import csv
import json
import os
import time
import uuid
from datetime import datetime
from collections import defaultdict
from pprint import pprint
from terminaltables import SingleTable
from docxtpl import DocxTemplate
from common import create_doc_final

FOLDER_PATH = r'logs' #r'Alarms'  # r'\\192.168.25.97\c\Logs\Scale\Trunk_conveyor'
# FILENAME = 'Trunk_scale_weight_hour_log_0.csv'
FILENAME = 'Trunk_scale_weight_minute_log_0.csv'
TEMPLATE_PATH = 'шаблон.docx'


YEAR = 2023
MONTH = 3
DAY = 1


# def get_events_table(events):
#     sorted_events = sorted(events.items(), key=lambda i: i[0])
#     table_data = []
#     for event in sorted_events:
#         times = ''
#         more_events_indicator = ' '
#         if len(event[1]) > EVENTS_QTY:
#             more_events_indicator = f'...({len(event[1])} events)'
#         for item in event[1][:EVENTS_QTY]:
#             times += f" {item['log_row_time']}"
#         table_data.append(
#                 [datetime.strftime(event[0], "%d %B %Y"), f'{times}{more_events_indicator}']
#             )
#     table_instance = SingleTable(table_data)
#     table_instance.inner_heading_row_border = False
#     return table_instance.table


if __name__ == '__main__':
    parsed_log = []
    with open(os.path.join(FOLDER_PATH, FILENAME), 'r', newline='') as file:
        reader = csv.DictReader(file, delimiter=';')
        for row in reader:
            try:
                log_row_datetime = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S')
                log_row_date = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S').date()
                log_row_weight = round(float(row['VarValue'].replace(',', '.')), 3)
                if log_row_date.year == YEAR:
                    parsed_log.append({
                        'log_row_date': log_row_date,
                        'log_row_time': log_row_datetime,
                        'log_row_weight': log_row_weight
                    })
            except ValueError as ex:
                print(ex)
                continue

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

    context = {
        'report': report,
        'year_weight_sum': round(year_weight_sum, 3)
    }
    try:
        create_doc_final(TEMPLATE_PATH, context, 'scales_report')
    except PermissionError:
        create_doc_final(TEMPLATE_PATH, context, 'scales_report', exists=True)

    with open('scales_report.json', 'w', encoding='utf8') as file:
        json.dump(context, file, ensure_ascii=False)
