import glob
import os
from time import sleep
import getpass
from connection import hiv,sql,oracle
import pandas as pd
from selenium.webdriver import ActionChains
import pyodbc
import numpy as np
import numpy.testing as npt
import pyautogui
from selenium.webdriver.common.keys import Keys

input = pd.ExcelFile("Input.xlsx")
report = input.parse("reportname")
reportname = report.at[0, 'Report Name']
db = input.parse("database")
dbtype = db.at[0, 'Database']
user = getpass.getuser()

path = "http://acitgdprwn01:9300/bi/?perspective=home"


def validate(self):
    input = pd.ExcelFile("Input.xlsx")
    report = input.parse("reportname")
    reportname = report.at[0, 'Report Name']
    n = self.driver.find_element_by_xpath('//*[@id="com.ibm.bi.contentApps.myContentFoldersSlideout"]')
    n.click()
    sleep(10)
    # self.driver.find_element_by_xpath('//*[contains(text(),' +str(reportname)+')]').click()
    self.driver.find_element_by_xpath("//*[contains(text(), 'Order Detail Report')]").click()
    sleep(100)
    self.driver.find_element_by_xpath('//*[@id="com.ibm.bi.authoring.runBtn.menu"]').click()
    sleep(5)
    self.driver.find_element_by_xpath("//*[contains(text(), 'Run Excel')]").click()
    sleep(20)


    #Database report testing

    if dbtype == "Hive":
        hiv(self)
    elif dbtype == "Sql":
        sql(self)
    elif dbtype == "Oracle":
        oracle(self)
    else:

        print("Invalid database")
        os.remove(glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0])
        exit()



        # # comparing the 2 excel sheets
    # df_OLD = pd.read_excel("actuals.xlsx").fillna(0)
    #
    # try:
    #     path_NEW = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    # except IndexError:
    #     raise IOError("No .xlsx files found in ")
    # finally:
    #     df_NEW = pd.read_excel(path_NEW).fillna(0)
    #
    # dfDiff = df_OLD.copy()
    # error_count = 0
    # for row in range(dfDiff.shape[0]):
    #     for col in range(dfDiff.shape[1]):
    #         value_OLD = df_OLD.iloc[row, col]
    #         try:
    #             value_NEW = df_NEW.iloc[row, col]
    #             if value_OLD == value_NEW:
    #                 dfDiff.iloc[row, col] = df_NEW.iloc[row, col]
    #             else:
    #                 # dfDiff.iloc[row, col] = ('{}-->{}').format(value_OLD, value_NEW)
    #                 dfDiff.iloc[row, col] = "Error"
    #                 error_count = error_count + 1
    #         except:
    #             dfDiff.iloc[row, col] = "NaN"
    #
    # fname = "excel_diff.xlsx"
    # writer = pd.ExcelWriter(fname, engine='xlsxwriter')
    #
    # dfDiff.to_excel(writer, sheet_name='DIFF', index=False)
    # df_NEW.to_excel(writer, sheet_name="report data", index=False)
    # df_OLD.to_excel(writer, sheet_name="actual data", index=False)
    #
    # workbook = writer.book
    # worksheet = writer.sheets['DIFF']
    # worksheet.hide_gridlines(2)
    #
    # # define formats
    # grey_fmt = workbook.add_format({'font_color': '#030303'})
    # highlight_fmt = workbook.add_format({'font_color': '#e61515', 'bg_color': '#e61515'})
    #
    # ## highlight changed cells
    # worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
    #                                            'criteria': 'containing',
    #                                            'value': '→',
    #                                            'format': highlight_fmt})
    # ## highlight unchanged cells
    # worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
    #                                            'criteria': 'not containing',
    #                                            'value': '→',
    #                                            'format': grey_fmt})
    # # save
    # writer.save()
    #
    # if error_count > 0:
    #     print("Report Testing FAILED")
    # else:
    #     print("Report Testing PASSED")
    #
    # self.assertEqual(0, error_count, "Report Testing FAILED")
    #
    # os.remove(glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0])
    # # exit()
