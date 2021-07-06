
from openpyxl.chart import (
    PieChart,
    Reference
)
from openpyxl.chart.label import DataLabelList

def chart1(worksheet):
    pie = PieChart()
    pie.width = 12
    pie.height = 6

    labels = Reference(worksheet, min_col=1, min_row=7, max_row=9)
    data = Reference(worksheet, min_col=2, min_row=6, max_row=9)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Test Summary Result - Chrome"

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    worksheet.add_chart(pie, "D1")

def chart2(worksheet):
    pie = PieChart()
    pie.width = 12
    pie.height = 6

    labels = Reference(worksheet, min_col=1, min_row=12, max_row=14)
    data = Reference(worksheet, min_col=2, min_row=11, max_row=14)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Test Summary Result - Firefox"

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    worksheet.add_chart(pie, "D15")

def create_chart(worksheet):
    chart1(worksheet)
    chart2(worksheet)
