import openpyxl
from openpyxl.styles import PatternFill, Color, Font, Alignment
from openpyxl.utils import get_column_letter

errorFill = PatternFill(patternType='solid', fgColor='EE1111')
successFill = PatternFill(patternType='solid', fgColor='00AA00')
titleFill = PatternFill(patternType='solid', fgColor='68A0F9')
darkFill = PatternFill(patternType='solid', fgColor='000000')

textAlignment = Alignment(horizontal='left', vertical='bottom', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)

textFont = Font(name='Calibri', size=11, bold=False, color='ffffff')

def format_summary_title(worksheet):
    total_length = 7
    for len in range(1, total_length + 1):
        _cell = worksheet.cell(row=len, column=1)
        _cell.fill = titleFill
        _cell.font = textFont

def format_test_details_title(worksheet):
    total_length = 4
    for len in range(1, total_length + 1):
        _cell = worksheet.cell(row=1, column=len)
        _cell.fill = darkFill
        _cell.font = textFont

def fit_column(worksheet):
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2.5)
        worksheet.column_dimensions[column].width = adjusted_width


def format_testdetails(worksheet):
    format_test_details_title(worksheet)
    fit_column(worksheet)

def format_testsummary(worksheet):
    format_summary_title(worksheet)
    fit_column(worksheet)
