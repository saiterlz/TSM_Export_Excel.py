from datetime import date

from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,
    Reference,
)
from openpyxl.chart.axis import DateAxis

wb = Workbook()
ws = wb.active

rows = [
    ['Date', 'Batch 1', 'Batch 2', 'Batch 3'],
    [date(2015,9, 1), 40, 30, 25],
    [date(2015,9, 2), 40, 25, 30],
    [date(2015,9, 3), 50, 30, 45],
    [date(2015,9, 4), 30, 25, 40],
    [date(2015,9, 5), 25, 35, 30],
    [date(2015,9, 6), 20, 40, 35],
]

for row in rows:
    ws.append(row)

c1 = LineChart()
c1.title = "Line Chart" # 图的标题
c1.style = 8  # 线条的style
c1.y_axis.title = 'Size'  # y坐标的标题
c1.x_axis.title = 'Date' # x坐标的标题
c1.x_axis.number_format = 'd-mmm'  # 规定日期格式  这是月,年格式
c1.x_axis.majorTimeUnit = "Months"  # 规定日期间隔 注意days；Months大写

data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=7)        # 图像的数据 起始行、起始列、终止行、终止列
c1.add_data(data, titles_from_data=True)


ws.add_chart(c1, "A10")


# # Style the lines
# s1 = c1.series[0]
# s1.marker.symbol = "triangle"
# s1.marker.graphicalProperties.solidFill = "FF0000" # Marker filling
# s1.marker.graphicalProperties.line.solidFill = "FF0000" # Marker outline
#
# s1.graphicalProperties.line.noFill = True
#
# s2 = c1.series[1]
# s2.graphicalProperties.line.solidFill = "00AAAA"
# s2.graphicalProperties.line.dashStyle = "sysDot"
# s2.graphicalProperties.line.width = 100050 # width in EMUs
#
# s2 = c1.series[2]
# s2.smooth = True # Make the line smooth

# ws.add_chart(c1, "A10")


wb.save("line.xlsx")
