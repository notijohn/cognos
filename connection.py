import pandas as pd
import xlrd
import numpy as np
import glob
import os
import getpass
import pyodbc


user = getpass.getuser()


def hiv(self, i, c):
    ip = pd.ExcelFile("Input.xlsx")
    hive = ip.parse("Hive")
    driver = str(hive.at[0,'Driver'])
    Host = str(hive.at[0,'Host'])
    Port = str(hive.at[0,'Port'])
    Authmech = str(hive.at[0,'AuthMech'])
    DSN = str(hive.at[0,'DSN'])
    uid = str(hive.at[0,'User ID'])
    pwd = str(hive.at[0,'Password'])
    db = str(hive.at[0,'Database'])
    hs = str(hive.at[0,'Hive Server Type'])

    pyodbc.autocommit = True
    #table ="orders"
    conn_str = (
    r'DRIVER='+driver+';'     		
	r'Host='+Host+';'
	r'Port='+Port+';'	
	r'AuthMech='+Authmech+';'	
	r'DSN='+DSN+';'
	r'UID='+uid+';'
	r'PWD='+pwd+';'
	r'Database='+db+';'
	r'HiveServerType='+hs+';'	
    )

    conn = pyodbc.connect(conn_str, autocommit=True)
    
    
    ag = pd.ExcelFile("agg.xlsx")
         
    df9 = ag.parse(str(i))
    c1 = str(df9.at[0, c])
        

    df6= pd.read_sql_query(c1,conn).round() #Running the query to get the desired data from database
              
    try:
        path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    except IndexError:
        raise IOError("No .xlsx files found in ") 
    finally:
        d2 = pd.read_excel(path1).round() #Fetching the downloaded chart data

    d3 = dict(d2.to_dict())
    d3 = df6.sort_values(by=[df6.columns[1]])
    d3=d3.dropna()
    d3 = d3.to_numpy()
    
        
    #d4 = d2[[s]]
    d4 = dict(df6.to_dict())

    d4 = d2.sort_values(by=[d2.columns[1]])
    #d4=d4.astype(float)
    d4=d4.to_numpy()
    
        
    #d4=d4.reset_index()
    #d4 = d4.astype(float)

    #assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')
    
    result = np.array_equal(d3,d4)
    os.remove(path1)
            
    if str(result) == "True":
        print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
    
    self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")


def sql(self, i, c):
    '''enter code for sql connection..Refer Hive conenction code..modify according to database'''
    ag = pd.ExcelFile("agg.xlsx")
            
    df9 = ag.parse(str(i))
    c1 = str(df9.at[0, c])
    
    

    df6= pd.read_sql_query(c1,conn).round()
    
            

    try:
        path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    except IndexError:
        raise IOError("No .xlsx files found in ")
    finally:
        d2 = pd.read_excel(path1).round()

    d3 = dict(d2.to_dict())
    d3 = df6.sort_values(by=[df6.columns[1]])
    d3=d3.dropna()
    d3 = d3.to_numpy()
        
    
    #d4 = d2[[s]]
    d4 = dict(df6.to_dict())

    d4 = d2.sort_values(by=[d2.columns[1]])
    #d4=d4.astype(float)
    d4=d4.to_numpy()
    
        
    #d4=d4.reset_index()
    #d4 = d4.astype(float)

   
       
        

    #assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')
    
    result = np.array_equal(d3,d4)
    os.remove(path1)
            
    if str(result) == "True":
        print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
    
    self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")


def oracle(self, i, c):
    '''enter code for oracle connection..Refer Hive conenction code..modify according to database'''
    ag = pd.ExcelFile("agg.xlsx")
            
    df9 = ag.parse(str(i))
    c1 = str(df9.at[0, c])
    
      

    df6= pd.read_sql_query(c1,conn).round()
    
            

    try:
        path1 = glob.glob(os.path.join('C:/Users/' + str(user) + '/Downloads', "*.xlsx"))[0]
    except IndexError:
        raise IOError("No .xlsx files found in ")
    finally:
        d2 = pd.read_excel(path1).round()

    d3 = dict(d2.to_dict())
    d3 = df6.sort_values(by=[df6.columns[1]])
    d3=d3.dropna()
    d3 = d3.to_numpy()
    
    print(d3)
    
    #d4 = d2[[s]]
    d4 = dict(df6.to_dict())

    d4 = d2.sort_values(by=[d2.columns[1]])
    #d4=d4.astype(float)
    d4=d4.to_numpy()
    
    print(d4)
    
    #d4=d4.reset_index()
    #d4 = d4.astype(float)

   
       
        

    #assert_frame_equal(d4,d3,check_dtype=False, check_index_type=False, check_column_type=False, check_frame_type=False, check_less_precise=False, check_names=False, by_blocks=False, check_exact=False, check_datetimelike_compat=False, check_categorical=False, check_like=True, obj='DataFrame')
    
    result = np.array_equal(d3,d4)
    os.remove(path1)
            
    if str(result) == "True":
        print("Sheet No: " + str(i) + ", Graph: " + str(c) + " PASSED")
    
    self.assertEqual('True', str(result), "Sheet No: " + str(i) + ", Graph: " + str(c) + " Failed")
    




