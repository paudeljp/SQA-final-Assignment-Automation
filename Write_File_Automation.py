import openpyxl
import os
# import time
from datetime import datetime
time = str(datetime.now())

test_result_location = 'Output_Result/test_result/TestResult.xlsx'
test_summary_location = 'Output_Result/test_result/TestSummmary.xlsx'

def excel_creater():
    if(os.path.exists(test_result_location)):
        workbook = openpyxl.load_workbook(test_result_location)
        worksheet = workbook['TestResults']
        return workbook,worksheet
    else:
        workbook = openpyxl.Workbook()
        worksheet = workbook.create_sheet('TestResults')
        worksheet.cell(row=1, column=1).value = "SN"
        worksheet.cell(row=1, column=2).value = "Test Summary"
        worksheet.cell(row=1, column=3).value = "Result"
        worksheet.cell(row=1, column=4).value = "Remarks"
        workbook.save(test_result_location)
        return workbook,worksheet

def write_result(sn,test_summary,result,remarks):
    workbook,worksheet = excel_creater()
    fieldnames = (int(sn),test_summary,result,str(remarks))
    start_column = 1
    start_row = int(sn)+1
    for field in fieldnames:
        worksheet.cell(row=start_row,column=start_column).value = field
        start_column+=1
    workbook.save(test_result_location)

def write_summary():
    workbook = openpyxl.Workbook()
    worksheet = workbook.create_sheet('TestSummary')
    worksheet.cell(row=1, column=1).value = "Test Executed On"
    worksheet.cell(row=1, column=2).value = time
    worksheet.cell(row=2, column=1).value = "Number of Test Cases"
    worksheet.cell(row=2, column=2).value = "=COUNTA(Sheet1!A:A)"
    workbook.save(test_summary_location)

def remove_file():
    if (os.path.exists(test_result_location)):
        os.remove(test_result_location)

    if (os.path.exists(test_summary_location)):
        os.remove(test_summary_location)


