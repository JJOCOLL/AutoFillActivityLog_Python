# !/usr/bin/env python
# coding=utf-8

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time


def gotoPath(driver, path):
    driver.get(path)


def clickOn(driver, locator):
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, locator)))
        driver.find_element_by_css_selector(locator).click()
    except NoSuchElementException:
        print("No Ellement is found!")


def sendKey(driver, locator, keys):
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, locator)))
        driver.find_element_by_css_selector(locator).send_keys(keys)
    except NoSuchElementException:
        print("No Ellement is found!")


def loginSys(driver, name, key):
    clickOn(driver, "i[class='icon-signin']")

    sendKey(driver, "input[id='i0116']", name)
    clickOn(driver, 'input[id="idSIButton9"]')

    sendKey(driver, "input[id='i0118']", key)
    clickOn(driver, 'input[value="Sign in"]')

    clickOn(driver, 'input[id="idSIButton9"]')


def gotoActivityLog(driver):
    addActivityLog_path = "https://pms.cs.cityu.edu.hk/pms/student/addlog.jsp"
    gotoPath(driver, addActivityLog_path)


def type_detect(type):
    return {
        'Customer Training': '29',
        'Database': '23',
        'Documentation': '26',
        'Internet/Intranet Programming': '22',
        'Marketing Support Activities': '28',
        'Meeting': '32',
        'Multimedia': '24',
        'Network Administration': '21',
        'Network Support': '20',
        'Program Development': '19',
        'Project Management': '30',
        'Software Maintenance': '18',
        'Software/System Testing': '25',
        'System Study & Design': '17',
        'System Support': '31',
        'User Support': '27',
    }.get(type, '26')


def create_DailyLog(driver, man_days, type, activity, CILO1, CILO2):
    gotoActivityLog(driver)
    clickOn(driver, "div[id='datetimepicker1'] span[class='add-on']")
    clickOn(driver, "div[id='datetimepicker1'] span[class='add-on']")
    sendKey(driver, "input[name='Duration1']", str(man_days))
    clickOn(driver, "select[name=\'TypeID1\'] option[value=\'" + type_detect(type) + "\']")
    sendKey(driver, "textarea[name='Activity1']", activity)
    clickOn(driver, "select[name='Progress1'] option[value='1']")
    clickOn(driver, "button[class='multiselect dropdown-toggle btn btn-default']")
    driver.find_element_by_xpath(
        '/html/body/div[1]/div[3]/form/table/tbody/tr/td[7]/div/ul/li[' + str(CILO1) + ']').click()
    driver.find_element_by_xpath(
        '/html/body/div[1]/div[3]/form/table/tbody/tr/td[7]/div/ul/li[' + str(CILO2) + ']').click()
    clickOn(driver, "button[id='add']")


def main():
    startTime = time.perf_counter()
    print('Start to fill activity log')

    Options.binary_location = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    options = Options()
    webDriver_path = r"\\corp.ha.org.hk\ha\HO\FD\IT\TOSS\T4\08Others\Justin liang\Task\chromedriver.exe"
    driver = webdriver.Chrome(executable_path=webDriver_path, options=options)

    mainPage_path = "https://pms.cs.cityu.edu.hk/pms/"
    gotoPath(driver, mainPage_path)

    # driver.set_window_position(-3000, 0)  # driver.set_window_position(0, 0) to get it back

    loginName = ""
    key = ""
    loginSys(driver, loginName, key)

    # Documentation log
    create_DailyLog(driver, 0.3, 'Documentation', 'Pager record ; HUS DP Pool Utilization ; Daily change', 1, 6)

    # System Support log
    create_DailyLog(driver, 0.2, 'System Support', 'Basic check for hosts, Final check for hosts', 2, 5)

    # Program Development log
    create_DailyLog(driver, 0.3, 'Program Development', 'Development of python and vba programs', 3, 4)
    print('Automation of filling activity log is successful!')

    endTime = time.perf_counter() - startTime
    print('Total time spent: {:.2f}s'.format(endTime))

    # time.sleep(60)
    driver.close()


if __name__ == "__main__":
    main()
