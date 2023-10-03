import csv
from datetime import datetime
from collections import defaultdict

from terminaltables import SingleTable

FOLDER_PATH = r'\\192.168.25.97\c\Logs\Alarms' #r'Alarms'  # r'\\192.168.25.97\c\Logs\Alarms'
FILES_COUNT = 11

ERROR_TXT = 'Обрыв связи с S7-1200 ШУ2'
YEAR = 2023
MONTH = 7
DAY = 1


def table_events(events):
    sorted_events = sorted(events.items(), key=lambda i: i[0])
    table_data = []
    for event in sorted_events:
        times = ''
        for item in event[1]:
            times += f" {item['log_row_time']}"
        table_data.append(
                [datetime.strftime(event[0], "%d %B %Y"), times]
            )
    table_instance = SingleTable(table_data)
    table_instance.inner_heading_row_border = False
    return table_instance.table


if __name__ == '__main__':
    parsed_log = []
    target_datetime = datetime(YEAR, MONTH, DAY)
    for idx in range(FILES_COUNT):
        with open(f'{FOLDER_PATH}\Alarm_Log_{idx}.csv', 'r', newline='') as file:
            reader = csv.DictReader(file, delimiter=';')
            for row in reader:
                try:
                    log_row_datetime = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S')
                    log_row_date = datetime.strptime(row['TimeString'], '%d.%m.%Y %H:%M:%S').date()
                    log_row_msg = row['MsgText']

                    if log_row_msg == ERROR_TXT and log_row_datetime > target_datetime:
                        parsed_log.append({
                            'log_row_date': log_row_date,
                            'log_row_time': datetime.strftime(log_row_datetime, '%H:%M:%S'),
                        })
                except ValueError as ex:
                    #print(ex)
                    continue
    events = defaultdict(list)
    for event in parsed_log:
        events[event['log_row_date']].append(event)
    print(table_events(events))
