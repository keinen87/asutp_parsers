import json
from pprint import pprint
from common import create_doc_final

TEMPLATE_PATH = 'шаблон для сравнения.docx'


if __name__ == '__main__':
    with open('scales_report.json', 'r', encoding='utf8') as file:
        scales_report = json.load(file)

    with open('disp_report.json', 'r', encoding='utf8') as file:
        disp_report = json.load(file)
    year_weight_sum = scales_report['year_weight_sum'] - disp_report['year_weight_sum']   
    report = []

    for s_report, d_report in zip(scales_report, disp_report): 
        try:
            for s_month_info, d_month_info in zip(scales_report[s_report], disp_report[d_report]):
                s_month_weight_sum = round(list(s_month_info.items())[0][1]['month_weight_sum'], 2)
                d_month_weight_sum = round(list(d_month_info.items())[0][1]['month_weight_sum'], 2)
                current_month = list(s_month_info.keys())[0]
                print(current_month)
                if s_month_weight_sum == 0 and d_month_weight_sum == 0:
                    month_report = {
                        current_month: {
                            'month_weight_delta': 0,
                            'days': []
                        }
                    }
                    report.append(month_report)
                    continue
                elif d_month_weight_sum == 0:
                    month_weight_delta = s_month_weight_sum - 0
                    month_report = {
                        current_month: {
                            'month_weight_delta': round(month_weight_delta, 2),
                            'days': []
                        }
                    }
                    report.append(month_report)
                    continue
                else:
                    month_weight_delta = s_month_weight_sum - d_month_weight_sum
                    s_current_month_days = list(s_month_info.items())[0][1]['days']
                    d_current_month_days = list(d_month_info.items())[0][1]['days']
                    first_day_scale_cur_month = list(s_current_month_days[0].keys())[0]
                    first_day_disp_cur_month = list(d_current_month_days[0].keys())[0]
                    while first_day_scale_cur_month != first_day_disp_cur_month:
                        d_current_month_days.pop(0)
                        first_day_disp_cur_month = list(d_current_month_days[0].keys())[0]
                    d_current_month_days = d_current_month_days[:len(s_current_month_days)]
                    days = []
                    for s_day_info, d_day_info in zip(s_current_month_days, d_current_month_days):
                        day_num = list(s_day_info.keys())[0]
                        weight_delta = round(list(s_day_info.items())[0][1] - list(d_day_info.items())[0][1],2)
                        day = {
                            day_num: weight_delta
                        }
                        days.append(day)
                    month_report = {
                        current_month: {
                            'month_weight_delta': round(month_weight_delta, 2),
                            'days': days
                        }
                    }
                report.append(month_report)    
        except TypeError:
            break

    context = {
        'report': report,
        'year_weight_sum': round(year_weight_sum, 2)
    }
    try:
        create_doc_final(TEMPLATE_PATH, context, 'compare_report')
    except PermissionError:
        create_doc_final(TEMPLATE_PATH, context, 'compare_report', exists=True)