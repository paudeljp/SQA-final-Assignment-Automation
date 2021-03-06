from selenium import webdriver
import pandas as pd
import Write_File_Automation
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
import sendEmail

start_time = str(datetime.now())
url_name = 'https://webmd.com'

def read_excel():
    print("Selenium Web Automation Framework")
    reader = pd.read_excel('./Test_Case/TestCase.xlsx')
    for row,column in reader.iterrows():
        sn = column['SN']
        execute_flag = column['Execute_FLAG']
        test_summary = column['Test Summary']
        xpath = column['Xpath']
        action = column['Action']
        value = column['Value']
        if execute_flag != 'N':
            action_defination(sn,test_summary,xpath,action,value)
        else:
            result = 'SKIPPED'
            remarks = 'Test was skipped due to N FLAG'
            print(result,remarks)
            Write_File_Automation.write_result(sn, test_summary, result, driverValue, remarks)

def action_defination(sn,test_summary,xpath,action,value):
    try:
        if action == 'open_browser':
            open_browser_function(value)
            result = 'PASS'
            remarks = ''
        elif action == 'open_url':
            result,remarks = open_url_function(driver,value)
        elif action == 'openUrlInNew_tab':
            result,remarks = openUrlInNew_tab_function(driver,value,xpath)
        elif action == 'click':
            result,remarks = click_function(driver,xpath)
        elif action == 'verify_text':
            result,remarks = verify_text_function(driver,xpath)
        elif action == 'verify_title':
            result,remarks = verify_title_function(driver,value)
        elif action == 'send_value':
            result,remarks = send_value_function(driver,xpath,value)
        elif action == 'select_dropdown':
            result,remarks = select_dropdown_function(driver,xpath,value)
        elif action == 'wait':
            result,remarks = wait_function(value)
        elif action == 'hover':
            result,remarks = hover_function(driver,xpath)
        elif action == 'close_browser':
            result,remarks = close_browser_function(driver)
        else:
            result = 'FAIL'
            remarks = 'Action defination not found'
            print(result,remarks)
        print(sn,test_summary,result,remarks)
        Write_File_Automation.write_result(sn,test_summary,result, driverValue, remarks)
    except Exception as ex:
        print("Exception has occured")
        result = 'FAIL'
        remarks = ex
        print(result,remarks)
        Write_File_Automation.write_result(sn, test_summary, result, driverValue, remarks)

def open_browser_function(value):
    global driver
    global driverValue
    driverValue = value
    if value == 'Chrome':
        print('Chrome Browser Selected')
        driver = webdriver.Chrome('Driver_Path/chromedriver.exe')
        driver.maximize_window()
    elif value == 'Firefox':
        print('Firefox Browser Selected')
        driver = webdriver.Firefox(executable_path="Driver_Path/firefox/geckodriver.exe")
    elif value == 'Safari':
        driver = webdriver.Safari()
    else:
        print("Browser not supported")
    return driver

def open_url_function(driver,value):
    try:
        driver.get(value)
        result = 'PASS'
        remarks = ""
    except Exception as ex:
        result = 'FAIL'
        remarks = ex
    return result,remarks

def openUrlInNew_tab_function(driver,value,xpath):
    try:
        find_doctor_xpath = driver.find_element_by_xpath(xpath)
        action = ActionChains(driver)
        action.key_down(Keys.CONTROL).click(find_doctor_xpath).key_up(Keys.CONTROL).perform()
        driver.switch_to.window(driver.window_handles[0])
        driver.get(value)
        result = 'PASS'
        remarks = ''
    except Exception as ex:
        result = 'FAIL'
        remarks = ex
    return result,remarks


def click_function(driver,xpath):
    try:
        driver.find_element_by_xpath(xpath).click()
        result,remarks = 'PASS', ''
    except Exception as ex:
        result = 'FAIL'
        remarks = ex
    return result,remarks

def verify_text_function(driver,xpath):
    try:
        driver.find_element_by_xpath(xpath)
        result,remarks = 'PASS',''
    except Exception as ex:
        result,remarks = 'FAIL', ex
    return result,remarks

def send_value_function(driver,xpath,value):
    try:
        driver.find_element_by_xpath(xpath).send_keys(value)
        result,remarks = 'PASS', ''
    except Exception as ex:
        result,remarks = 'FAIL', ex
    return result,remarks

def select_dropdown_function(driver,xpath,value):
    try:
        driver.find_element_by_xpath(xpath).send_keys(value)
        result,remarks = 'PASS', ''
    except Exception as ex:
        result,remarks = 'FAIL', ex
    return result,remarks

def wait_function(value):
    try:
        time.sleep(value)
        result,remarks = 'PASS', ''
    except Exception as ex:
        result,remarks = 'FAIL', ex
    return result,remarks

def verify_title_function(driver ,value):
    actual_text = driver.title
    expected_text = value
    try:
        assert expected_text == actual_text
    except AssertionError:
        result,remarks = 'FAIL', 'Title not matched'
    else:
        result,remarks = 'PASS', ''
    return result,remarks

def hover_function(driver,xpath):
    action = ActionChains(driver)
    try:
        hover_element = driver.find_element_by_xpath(xpath)
        action.move_to_element(hover_element).perform()
        result,remarks = 'PASS', ''
    except Exception as ex:
        result,remarks = 'FAIL', ex
    return result,remarks

def close_browser_function(driver):
    try:
        driver.quit()
        result,remarks = 'PASS', ''
    except Exception as ex:
        result,remarks = 'FAIL',ex
    return result,remarks


if __name__ == "__main__":
    Write_File_Automation.remove_file()
    print("file removed")
    read_excel()
    Write_File_Automation.write_summary(start_time, url_name)
    Write_File_Automation.format_excelsheet()
    sendEmail.send_report()

