import unittest
from selenium import webdriver
from time import sleep
from validation import validate
from highlight_ele import highlight
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlrd
import glob
import os
from selenium.webdriver.common.alert import Alert
import time
import win32com.client
import keyboard
import pywinauto
import pyautogui

input = pd.ExcelFile("Input.xlsx")
report = input.parse("reportname")
reportname = report.at[0,'Report Name']

path = "http://acitgdprwn01:9300/bi/?perspective=home"


class SearchText(unittest.TestCase):

    @classmethod
    def setUpClass(self):

        # create new chrome session
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(50)
        self.driver.maximize_window()
        

        # navigate to the application home page
        self.driver.get(path)
        time.sleep(10)

    def test_sheet(self):
        validate(self)

    @classmethod
    def tearDownClass(self):

        self.driver.quit()
