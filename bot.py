#! /usr/lib/python3.6
# ms_rewards.py - Searches for results via pc bing browser and mobile, completes quizzes on pc bing browser
# Version 2019.07.13

# TODO replace sleeps with minimum sleeps for explicit waits to work, especially after a page redirect
# FIXME mobile version does not require re-sign in, but pc version does, why?
# FIXME Known Cosmetic Issue - logged point total caps out at the point cost of the item on wishlist

import argparse
import json
import logging
import os
import platform
import random
import time
import zipfile
import os
from datetime import datetime, timedelta

import requests
from requests.exceptions import RequestException
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException, \
    ElementClickInterceptedException, ElementNotVisibleException, \
    ElementNotInteractableException, NoSuchElementException, UnexpectedAlertPresentException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

# URLs
BING_SEARCH_URL = 'https://www.bing.com/search'
DASHBOARD_URL = 'https://account.microsoft.com/rewards/dashboard'
POINT_TOTAL_URL = 'http://www.bing.com/rewardsapp/bepflyoutpage?style=chromeextension'

# user agents for edge/pc and mobile
PC_USER_AGENT = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                 'AppleWebKit/537.36 (KHTML, like Gecko) '
                 'Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134')
MOBILE_USER_AGENT = ('Mozilla/5.0 (Windows Phone 10.0; Android 4.2.1; WebView/3.0) '
                     'AppleWebKit/537.36 (KHTML, like Gecko) coc_coc_browser/64.118.222 '
                     'Chrome/52.0.2743.116 Mobile Safari/537.36 Edge/15.15063')
# log levels
_LOG_LEVEL_STRINGS = ['CRITICAL', 'ERROR', 'WARNING', 'INFO', 'DEBUG']


def check_python_version():
    """
    Ensure the correct version of Python is being used.
    """
    minimum_version = ('3', '6')
    if platform.python_version_tuple() < minimum_version:
        message = 'Only Python %s.%s and above is supported.' % minimum_version
        raise Exception(message)

def _log_level_string_to_int(log_level_string):
    log_level_string = log_level_string.upper()

    if log_level_string not in _LOG_LEVEL_STRINGS:
        message = f'invalid choice: {log_level_string} (choose from {_LOG_LEVEL_STRINGS})'
        raise argparse.ArgumentTypeError(message)

    log_level_int = getattr(logging, log_level_string, logging.INFO)
    # check the logging log_level_choices have not changed from our expected values
    assert isinstance(log_level_int, int)
    return log_level_int


def init_logging(log_level):
    # gets dir path of python script, not cwd, for execution on cron
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    os.makedirs('logs', exist_ok=True)
    log_path = os.path.join('logs', 'ms_rewards.log')
    logging.basicConfig(
        filename=log_path,
        level=log_level,
        format='%(asctime)s :: %(levelname)s :: %(name)s :: %(message)s')

def browser_setup(headless_mode, user_agent):
    """
    Inits the chrome browser with headless setting and user agent
    :param headless_mode: Boolean
    :param user_agent: String
    :return: webdriver obj
    """
    os.makedirs('drivers', exist_ok=True)
    path = os.path.join('drivers', 'chromedriver')
    system = platform.system()
    if system == "Windows":
        if not path.endswith(".exe"):
            path += ".exe"
    if not os.path.exists(path):
        download_driver(path, system)

    options = Options()
    options.add_argument(f'user-agent={user_agent}')
    options.add_argument('--disable-webgl')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_experimental_option('w3c', False)
    options.add_extension('extensions/sourceScrub.crx')
    prefs = {
        "profile.default_content_setting_values.geolocation" : 2, "profile.default_content_setting_values.notifications": 2
        }

    options.add_experimental_option("prefs", prefs)

    if headless_mode:
        options.add_argument('--headless')

    chrome_obj = webdriver.Chrome(path, options=options)

    return chrome_obj


def find_by_class(selector):
    """
    Finds elements by class name
    :param selector: Class selector of html obj
    :return: returns a list of all matching selenium objects
    """
    return browser.find_elements_by_class_name(selector)


def find_by_css(selector):
    """
    Finds nodes by css selector
    :param selector: CSS selector of html node obj
    :return: returns a list of all matching selenium objects
    """
    return browser.find_elements_by_css_selector(selector)


def wait_until_visible(by_, selector, time_to_wait=10):
    """
    Searches for selector and if found, end the loop
    Else, keep repeating every 2 seconds until time elapsed, then refresh page
    :param by_: string which tag to search by
    :param selector: string selector
    :param time_to_wait: int time to wait
    :return: Boolean if selector is found
    """
    start_time = time.time()
    while (time.time() - start_time) < time_to_wait:
        if browser.find_elements(by=by_, value=selector):
            return True
        browser.refresh()  # for other checks besides points url
        time.sleep(2)
    return False


def wait_until_clickable(by_, selector, time_to_wait=10):
    """
    Waits 5 seconds for element to be clickable
    :param by_:  BY module args to pick a selector
    :param selector: string of xpath, css_selector or other
    :param time_to_wait: Int time to wait
    :return: None
    """
    try:
        WebDriverWait(browser, time_to_wait).until(ec.element_to_be_clickable((by_, selector)))
    except TimeoutException:
        logging.exception(msg=f'{selector} element Not clickable - Timeout Exception', exc_info=False)
        screenshot(selector)
    except UnexpectedAlertPresentException:
        # FIXME
        browser.switch_to.alert.dismiss()
        # logging.exception(msg=f'{selector} element Not Visible - Unexpected Alert Exception', exc_info=False)
        # screenshot(selector)
        # browser.refresh()
    except WebDriverException:
        logging.exception(msg=f'Webdriver Error for {selector} object')
        screenshot(selector)


def send_key_by_name(name, key):
    """
    Sends key to target found by name
    :param name: Name attribute of html object
    :param key: Key to be sent to that object
    :return: None
    """
    try:
        browser.find_element_by_name(name).send_keys(key)
    except (ElementNotVisibleException, ElementClickInterceptedException, ElementNotInteractableException):
        logging.exception(msg=f'Send key by name to {name} element not visible or clickable.')
    except NoSuchElementException:
        logging.exception(msg=f'Send key to {name} element, no such element.')
        screenshot(name)
        browser.refresh()
    except WebDriverException:
        logging.exception(msg=f'Webdriver Error for send key to {name} object')


def send_key_by_id(obj_id, key):
    """
    Sends key to target found by id
    :param obj_id: ID attribute of the html object
    :param key: Key to be sent to that object
    :return: None
    """
    try:
        browser.find_element_by_id(obj_id).send_keys(key)
    except (ElementNotVisibleException, ElementClickInterceptedException, ElementNotInteractableException):
        logging.exception(msg=f'Send key by ID to {obj_id} element not visible or clickable.')
    except NoSuchElementException:
        logging.exception(msg=f'Send key by ID to {obj_id} element, no such element')
        screenshot(obj_id)
        browser.refresh()
    except WebDriverException:
        logging.exception(msg=f'Webdriver Error for send key by ID to {obj_id} object')


def click_by_class(selector):
    """
    Clicks on node object selected by class name
    :param selector: class attribute
    :return: None
    """
    try:
        browser.find_element_by_class_name(selector).click()
    except (ElementNotVisibleException, ElementClickInterceptedException, ElementNotInteractableException):
        logging.exception(msg=f'Send key by class to {selector} element not visible or clickable.')
    except WebDriverException:
        logging.exception(msg=f'Webdriver Error for send key by class to {selector} object')


def click_by_id(obj_id):
    """
    Clicks on object located by ID
    :param obj_id: id tag of html object
    :return: None
    """
    try:
        browser.find_element_by_id(obj_id).click()
    except (ElementNotVisibleException, ElementClickInterceptedException, ElementNotInteractableException):
        logging.exception(msg=f'Click by ID to {obj_id} element not visible or clickable.')
    except WebDriverException:
        logging.exception(msg=f'Webdriver Error for click by ID to {obj_id} object')


def clear_by_id(obj_id):
    """
    Clear object found by id
    :param obj_id: ID attribute of html object
    :return: None
    """
    try:
        browser.find_element_by_id(obj_id).clear()
    except (ElementNotVisibleException, ElementNotInteractableException):
        logging.exception(msg=f'Clear by ID to {obj_id} element not visible or clickable.')
    except NoSuchElementException:
        logging.exception(msg=f'Send key by ID to {obj_id} element, no such element')
        screenshot(obj_id)
        browser.refresh()
    except WebDriverException:
        logging.exception(msg='Error.')


def main_window():
    """
    Closes current window and switches focus back to main window
    :return: None
    """
    try:
        for i in range(1, len(browser.window_handles)):
            browser.switch_to.window(browser.window_handles[i])
            browser.close()
    except WebDriverException:
        logging.error('Error when switching to main_window')
    finally:
        browser.switch_to.window(browser.window_handles[0])

def screenshot(selector):
    """
    Snaps screenshot of webpage when error occurs
    :param selector: The name, ID, class, or other attribute of missing node object
    :return: None
    """
    logging.exception(msg=f'{selector} cannot be located.')
    screenshot_file_name = f'{datetime.now().strftime("%Y%m%d%%H%M%S")}_{selector}.png'
    screenshot_file_path = os.path.join('logs', screenshot_file_name)
    browser.save_screenshot(screenshot_file_path)


def latest_window():
    """
    Switches to newest open window
    :return:
    """
    browser.switch_to.window(browser.window_handles[-1])




def ensure_pc_mode_logged_in():
    """
    Navigates to www.bing.com and clicks on ribbon to ensure logged in
    PC mode for some reason sometimes does not fully recognize that the user is logged in
    :return: None
    """
    browser.get(BING_SEARCH_URL)
    time.sleep(0.1)
    # click on ribbon to ensure logged in
    wait_until_clickable(By.ID, 'id_l', 15)
    click_by_id('id_l')
    time.sleep(0.1)


if __name__ == '__main__':
    try:
     browser = browser_setup(False, PC_USER_AGENT)
     browser.get('https://www.capterra.com/payment-processing-software/')
     parentWindow = browser.window_handles[0]
     time.sleep(2)
     buttons = find_by_css(".Button__StyledA-sc-1p3sq94-1.hDtTdr")
     count = 1
     for i in range(len(buttons)):
        if i % 2 == 0:
            buttons[i].click()
            
            time.sleep(2)
            browser.switch_to.window(browser.window_handles[1])
                
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            buttons = find_by_css(".Button__StyledA-sc-1p3sq94-1.hDtTdr")
            count += 1
            if(count>9):
                browser.quit()
     
    except WebDriverException:
        logging.exception(msg='Failure at main()')
