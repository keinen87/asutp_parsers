import csv
import plotly.graph_objects as go
from datetime import datetime

filepath = r'\\192.168.25.97\c\Logs\Data\Currents\Trunk_conveyor'

fig = go.Figure(
            layout_title_text='График'
        )
parsed_log = []
for i in range(11):
    with open(f"{filepath}\Winch_Tension_Log_{i}.csv", 'r', newline='') as file:
        reader = csv.DictReader(file, delimiter=';')
        try: 
            for row in reader:
                parsed_log.append((datetime.strptime(row['TimeString'],"%d.%m.%Y %H:%M:%S"), float(row['VarValue'].replace(',','.'))))
        except ValueError as ex:
            print(row)

        
fig.add_trace(go.Scatter(
        x = [datetime for datetime, _ in sorted(parsed_log, key=lambda z: z[0])],
        y = [speedvalue for _, speedvalue in sorted(parsed_log, key=lambda z: z[0])]
    ))
fig.show()
