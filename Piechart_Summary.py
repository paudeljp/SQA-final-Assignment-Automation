from openpyxl import Workbook

from openpyxl.chart import (
    PieChart,
    ProjectedPieChart,
    Reference
)
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList


def create_chart(worksheet):

    pie = PieChart()
    labels = Reference(worksheet, min_col=1, min_row=5, max_row=7)
    data = Reference(worksheet, min_col=2, min_row=4, max_row=7)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Test Summary Result"

    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    worksheet.add_chart(pie, "D1")
