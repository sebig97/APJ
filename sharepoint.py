import time
import pyautogui
import sys

from Tools.scripts.make_ctype import values
from selenium import webdriver
import requests

# def move_to_sharepoint(self):
#     # driver = self.driver
#     pyautogui.keyDown('alt')
#     time.sleep(.2)
#     pyautogui.press('tab')
#     time.sleep(.2)
#     pyautogui.press('tab')
#     time.sleep(.2)
#     pyautogui.keyUp('alt')
#     # time.sleep(1)
#     # pyautogui.click()
#
#     # driver.execute_script("window.open('');")
#     # driver.switch_to.window(driver.window_handles[1])
#     # driver.get('https://hp.sharepoint.com/sites/CCRobotics/RocketAPJ%20Credit%20Card/Forms/AllItems.aspx')
#     # time.sleep(3)
#
#
# def rename_from_sharepoint(self):
#     driver = self.driver
#
#     driver.find_element_by_xpath("//span[@class='signalField_05013969 wide_05013969']").click()
#     driver.find_element_by_xpath("//span[@id='id__651']").click()
#     driver.find_element_by_xpath("//input[@id='TextField688']").send_keys("CSV_no_" + str(1))
#     driver.find_element_by_xpath("//span[@id='id__691']").click()
#
#     driver.switch_to.window(driver.window_handles[0])

import requests
from requests_ntlm import HttpNtlmAuth
import getpass

SITE = "https://sharepointsite.com/"
USERNAME = "domain\\user"

response = requests.get(SITE, auth=HttpNtlmAuth(USERNAME, getpass.getpass()))
print(response.status_code)
