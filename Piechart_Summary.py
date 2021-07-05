from openpyxl import Workbook

from openpyxl.chart import (
    PieChart,
    ProjectedPieChart,
    Reference
)
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList

def chart1(worksheet):
    pie = PieChart()
    labels = Reference(worksheet, min_col=1, min_row=5, max_row=7)
    data = Reference(worksheet, min_col=2, min_row=4, max_row=7)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Test Summary Result - Chrome"

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    worksheet.add_chart(pie, "D1")

def chart2(worksheet):
    pie = PieChart()
    labels = Reference(worksheet, min_col=1, min_row=8, max_row=10)
    data = Reference(worksheet, min_col=2, min_row=7, max_row=10)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Test Summary Result - Firefox"

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    worksheet.add_chart(pie, "D20")

def create_chart(worksheet):
    chart1(worksheet)
    chart2(worksheet)
