import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting import Rule
from openpyxl.styles.differential import DifferentialStyle

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

def format_test_details(worksheet):
    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00AA00")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="PASS", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("PASS",C1)))']
    worksheet.conditional_formatting.add('C1:C60', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="EE1111")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="FAIL", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("FAIL",C1)))']
    worksheet.conditional_formatting.add('C1:C60', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="68A0F9")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="SKIPPED", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("SKIPPED",C1)))']
    worksheet.conditional_formatting.add('C1:C60', rule)

def format_testdetails(worksheet):
    format_test_details_title(worksheet)
    format_test_details(worksheet)
    fit_column(worksheet)

def format_testsummary(worksheet):
    format_summary_title(worksheet)
    fit_column(worksheet)
