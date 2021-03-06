import openpyxl
import os
from datetime import datetime
import Format_ExcelSheet
import Piechart_Summary

test_result_location = 'Output_Result/test_result/TestResult.xlsx'

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
        worksheet.cell(row=1, column=4).value = "Tested On"
        worksheet.cell(row=1, column=5).value = "Remarks"

        workbook.save(test_result_location)
        return workbook,worksheet

def write_result(sn,test_summary,result,driverValue,remarks):
    workbook,worksheet = excel_creater()
    fieldnames = (int(sn),test_summary,result,driverValue,str(remarks))
    start_column = 1
    start_row = int(sn)+1
    for field in fieldnames:
        worksheet.cell(row=start_row,column=start_column).value = field
        start_column+=1
    workbook.save(test_result_location)

def rename_sheet():
    workbook = openpyxl.load_workbook(test_result_location)
    worksheet = workbook['Sheet']
    worksheet.title = 'TestSummary'
    workbook.save(test_result_location)

def write_summary(start_time, url_name):
    rename_sheet()
    end_time = str(datetime.now())
    workbook = openpyxl.load_workbook(test_result_location)
    worksheet = workbook['TestSummary']
    worksheet.cell(row=1, column=1).value = "Test Executed On"
    worksheet.cell(row=1, column=2).value = start_time

    worksheet.cell(row=2, column=1).value = "Test Completed On"
    worksheet.cell(row=2, column=2).value = end_time

    worksheet.cell(row=3, column=1).value = "URL"
    worksheet.cell(row=3, column=2).value = url_name

    worksheet.cell(row=4, column=1).value = "Total Number of Test"
    worksheet.cell(row=4, column=2).value = "=((COUNTA(TestResults!A:A) - 1) / 2)"

    worksheet.cell(row=6, column=1).value = "CHROME"

    worksheet.cell(row=7, column=1).value = "Number of Passed Test Case - Chrome"
    worksheet.cell(row=7, column=2).value = '=COUNTIFS(TestResults!C:C,"PASS", TestResults!D:D,"Chrome")'

    worksheet.cell(row=8, column=1).value = "Number of Failed Test Case - Chrome"
    worksheet.cell(row=8, column=2).value = '=COUNTIFS(TestResults!C:C,"FAIL", TestResults!D:D,"Chrome")'

    worksheet.cell(row=9, column=1).value = "Number of Skipped Test Case - Chrome"
    worksheet.cell(row=9, column=2).value = '=COUNTIFS(TestResults!C:C,"Skipped", TestResults!D:D,"Chrome")'

    worksheet.cell(row=11, column=1).value = "FIREFOX"

    worksheet.cell(row=12, column=1).value = "Number of Passed Test Case - Firefox"
    worksheet.cell(row=12, column=2).value = '=COUNTIFS(TestResults!C:C,"PASS", TestResults!D:D,"Firefox")'

    worksheet.cell(row=13, column=1).value = "Number of Failed Test Case - Firefox"
    worksheet.cell(row=13, column=2).value = '=COUNTIFS(TestResults!C:C,"FAIL", TestResults!D:D,"Firefox")'

    worksheet.cell(row=14, column=1).value = "Number of Skipped Test Case - Firefox"
    worksheet.cell(row=14, column=2).value = '=COUNTIFS(TestResults!C:C,"Skipped", TestResults!D:D,"Firefox")'

    worksheet.cell(row=16, column=1).value = "Test Prepared By"
    worksheet.cell(row=16, column=2).value = 'Jeevan Paudel'

    workbook.save(test_result_location)

def format_excelsheet():
    workbook = openpyxl.load_workbook(test_result_location)

    testResultworksheet = workbook['TestResults']
    Format_ExcelSheet.format_testdetails(testResultworksheet)

    testSummaryworksheet = workbook['TestSummary']
    Format_ExcelSheet.format_testsummary(testSummaryworksheet)
    Piechart_Summary.create_chart(testSummaryworksheet)

    workbook.save(test_result_location)

def remove_file():
    if (os.path.exists(test_result_location)):
        os.remove(test_result_location)
    else:
        print("The file does not exist")

