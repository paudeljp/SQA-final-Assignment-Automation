from selenium import webdriver
import pandas as pd
import Write_File_Automation
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time




def read_excel():
    print("Hello")
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
            result = 'Not Tested'
            remarks = 'Test was skipped due to N FLAG'
            print(result,remarks)

def action_defination(sn,test_summary,xpath,action,value):
    try:
        if action == 'open_browser':
            open_browser_function(value)
            result = 'PASS'
            remarks = ''
        elif action == 'open_url':
            result,remarks = open_url_function(driver,value)
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
        # Write_File_Automation.write_result(sn,test_summary,result,remarks)
        # excel_operation.write_data()
    except Exception as ex:
        print("Exception has occured")
        result = 'FAIL'
        remarks = ex
        print(result,remarks)
        # Write_File_Automation.write_result(sn, test_summary, result, remarks)

def open_browser_function(value):
    global driver
    if value == 'Chrome':
        print('Chrome Browser Selected')
        driver = webdriver.Chrome('Driver_Path/chromedriver.exe')
        driver.maximize_window()
    elif value == 'Firefox':
        driver = webdriver.Firefox()
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

    # output_text = driver.find_element_by_xpath(xpath).text
    # try:
    #     assert output_text == value
    # except AssertionError:
    #     result = 'FAIL'
    #     remarks = 'Actual value is ' + output_text + 'Input value is' + value
    # else:
    #     result = 'PASS'
    #     remarks = ''
    # return result,remarks


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

# output_text = driver.find_element_by_xpath(xpath).text
    # try:
    #     assert output_text == value
    # except AssertionError:
    #     result = 'FAIL'
    #     remarks = 'Actual value is ' + output_text + 'Input value is' + value
    # else:
    #     result = 'PASS'
    #     remarks = ''
    # return result,remarks
def hover_function(driver,xpath):
    action = ActionChains(driver)
    try:
        hover_element = driver.find_element_by_xpath(xpath)
        action.move_to_element(hover_element).perform()
        # driver.switch_to.window(driver.window_handles[0])
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
    # excel_operation.remove_file()
    # print("file removed")
    read_excel()
    # excel_operation.write_summary()