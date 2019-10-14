import pandas as pd
import xlrd
import numpy as np
import glob
import os
import getpass
import pyodbc
user = getpass.getuser()


def hiv(self):
    ip = pd.ExcelFile("Input.xlsx")
    hive = ip.parse("Hive")
    driver = str(hive.at[0, 'Driver'])
    Host = str(hive.at[0, 'Host'])
    Port = str(hive.at[0, 'Port'])
    Authmech = str(hive.at[0, 'AuthMech'])
    DSN = str(hive.at[0, 'DSN'])
    uid = str(hive.at[0, 'User ID'])
    pwd = str(hive.at[0, 'Password'])
    db = str(hive.at[0, 'Database'])
    hs = str(hive.at[0, 'Hive Server Type'])

    pyodbc.autocommit = True
    # table ="orders"
    conn_str = (
            r'DRIVER=' + driver + ';'
            r'Host=' + Host + ';'
            r'Port=' + Port + ';'
            r'AuthMech=' + Authmech + ';'
            r'DSN=' + DSN + ';'
            r'UID=' + uid + ';'
            r'PWD=' + pwd + ';'
            r'Database=' + db + ';'
            r'HiveServerType=' + hs + ';'
    )

    conn = pyodbc.connect(conn_str, autocommit=True)

    #ag = pd.ExcelFile("agg.xlsx")
    #
    #df9 = ag.parse(str(i))
    #c1 = str(df9.at[0, c])
    #
    #df6 = pd.read_sql_query(c1, conn).round()  # Running the query to get the desired data from database
    #
    #try:
    #    path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    #except IndexError:
    #    raise IOError("No .xlsx files found in ")
    #finally:
    #    d2 = pd.read_excel(path1).round()  # Fetching the downloaded chart data
    #
    #d3 = dict(d2.to_dict())
    #d3 = df6.sort_values(by=[df6.columns[1]])
    #d3 = d3.dropna()
    #d3 = d3.to_numpy()
    #
    ## d4 = d2[[s]]
    #d4 = dict(df6.to_dict())
    #
    #d4 = d2.sort_values(by=[d2.columns[1]])
    ## d4=d4.astype(float)
    #d4 = d4.to_numpy()

    # d4=d4.reset_index()
    # d4 = d4.astype(float)

    # assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')

    #result = np.array_equal(d3, d4)
    #os.remove(path1)
    #
    #if str(result) == "True":
    #    print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
    #
    #self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")


def sql(self):
    '''enter code for sql connection..Refer Hive conenction code..modify according to database'''

    def connectToSql(serverName, databaseName, execQuery, userName, password):
        try:
            cnxn = pyodbc.connect('Driver={ODBC Driver 11 for SQL Server};'
                                  'Server=' + serverName + ';'
                                  'Database=' + databaseName + ';'
                                  'username=' + userName + ';'
                                  'password=' + password + ';'
                                  'Trusted_Connection=yes;')
            #print(cnxn)
            #print(execQuery)
            df = pd.read_sql_query('' + execQuery + '', cnxn)
            #print(df)
            return df
        except Exception as e:
            return 0

    ip = pd.ExcelFile("Input.xlsx")
    sql = ip.parse("SQL")
    server = str(sql.at[0, 'Server'])
    uid = str(sql.at[0, 'Username'])
    pwd = str(sql.at[0, 'Password'])
    db = str(sql.at[0, 'Database'])
    query = str(sql.at[0, 'Query'])

    df_OLD = connectToSql(server, db, query, uid, pwd)
    # df_OLD = pd.read_excel("actuals.xlsx").fillna(0)
    try:
        path_NEW = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    except IndexError:
        raise IOError("No .xlsx files found in ")
    finally:
        df_NEW = pd.read_excel(path_NEW).fillna(0)

    dfDiff = df_OLD.copy()
    error_count = 0
    for row in range(dfDiff.shape[0]):
        for col in range(dfDiff.shape[1]):
            value_OLD = df_OLD.iloc[row, col]
            try:
                value_NEW = df_NEW.iloc[row, col]
                if value_OLD == value_NEW:
                    dfDiff.iloc[row, col] = df_NEW.iloc[row, col]
                else:
                    #dfDiff.iloc[row, col] = ('{}-->{}').format(value_OLD, value_NEW)
                    dfDiff.iloc[row, col] = "Error"
                    error_count = error_count + 1
            except:
                dfDiff.iloc[row, col] = "NaN"

    fname = "excel_diff.xlsx"
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='DIFF', index=False)
    df_NEW.to_excel(writer, sheet_name="report data", index=False)
    df_OLD.to_excel(writer, sheet_name="actual data", index=False)

    workbook = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)

    # define formats
    grey_fmt = workbook.add_format({'font_color': '#030303'})
    highlight_fmt = workbook.add_format({'font_color': '#e61515', 'bg_color': '#e61515'})

    ## highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                               'criteria': 'containing',
                                               'value': '→',
                                               'format': highlight_fmt})
    ## highlight unchanged cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                               'criteria': 'not containing',
                                               'value': '→',
                                               'format': grey_fmt})
    # save
    writer.save()

    if error_count > 0:
        print("Report Testing FAILED")

    self.assertEqual(0, error_count, "Report Testing PASSED")

# ag = pd.ExcelFile("agg.xlsx")
#
# df9 = ag.parse(str(i))
# c1 = str(df9.at[0, c])
#
# df6 = pd.read_sql_query(c1, conn).round()
#
# try:
#     path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
# except IndexError:
#     raise IOError("No .xlsx files found in ")
# finally:
#     d2 = pd.read_excel(path1).round()
#
# d3 = dict(d2.to_dict())
# d3 = df6.sort_values(by=[df6.columns[1]])
# d3 = d3.dropna()
# d3 = d3.to_numpy()
#
# # d4 = d2[[s]]
# d4 = dict(df6.to_dict())
#
# d4 = d2.sort_values(by=[d2.columns[1]])
# # d4=d4.astype(float)
# d4 = d4.to_numpy()
#
# # d4=d4.reset_index()
# # d4 = d4.astype(float)
#
#
# # assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')
#
# result = np.array_equal(d3, d4)
# os.remove(path1)
#
# if str(result) == "True":
#     print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
#
# self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")


def oracle(self):
    '''enter code for oracle connection..Refer Hive conenction code..modify according to database'''
    ag = pd.ExcelFile("agg.xlsx")
    #
    #df9 = ag.parse(str(i))
    #c1 = str(df9.at[0, c])
    #
    #df6 = pd.read_sql_query(c1, conn).round()
    #
    #try:
    #    path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    #except IndexError:
    #    raise IOError("No .xlsx files found in ")
    #finally:
    #    d2 = pd.read_excel(path1).round()
    #
    #d3 = dict(d2.to_dict())
    #d3 = df6.sort_values(by=[df6.columns[1]])
    #d3 = d3.dropna()
    #d3 = d3.to_numpy()
    #
    #print(d3)
    #
    ## d4 = d2[[s]]
    #d4 = dict(df6.to_dict())
    #
    #d4 = d2.sort_values(by=[d2.columns[1]])
    ## d4=d4.astype(float)
    #d4 = d4.to_numpy()
    #
    #print(d4)
    #
    ## d4=d4.reset_index()
    ## d4 = d4.astype(float)
    #
    ## assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')
    #
    #result = np.array_equal(d3, d4)
    #os.remove(path1)
    #
    #if str(result) == "True":
    #    print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
    #
    #self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")





