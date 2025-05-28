import pandas as pd
import os
from datetime import datetime, timedelta
import win32com.client
import pythoncom
import os
import pandas as pd
import shutil
import time
import json
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

config_file_path = r'\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\Config.csv'
limits_file_path = r'\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\Tech_Limits.csv'
# temp_file_path = 'Temp.csv'
server = 'rasinkul-desk'

def set_flags(x,limit_type,r_limit,y_limit,debug=False):
    flag = 'G'
    if limit_type == 'UCL':
        if x >= r_limit:
            flag = 'R'
        elif x>= y_limit:
            flag = 'Y'
    else:
        if x <= r_limit:
            flag = 'R'
        elif x<= y_limit:
            flag = 'Y'       

    if debug: print(f'        Type={limit_type}, Value={x}, R={r_limit}, Y={y_limit}, Flag={flag}') 
    return flag

def load_prod_data(server,product,details,limits,idv_unit='NONE',sicc_unit='NONE',debug=False):

    if debug: print(limits)
    tech = details['TECH']
    prod = details['PART']
    rev = details['REV']
    idv_unit = details['IDV_UNIT']
    sicc_unit = details['SICC_UNIT']
    cap_unit = details['CAPABILITY_UNIT']
    cdyn_unit = details['CDYN_UNIT']
    
    if debug: print(f'Product={product}')
    
    folders = r'\Actuals\Last_49_Days'
    if tech=='P1273':
        folders = r'\Actuals\Last_49Days'
    
    file_path = os.path.join(r'\\'+server,tech+r'_Data'+folders,prod+'.csv')
    
    # debug = True
    df = pd.read_csv(file_path)
    # if debug: print(df.head())

    # if debug:
    #     col_check = 1
    #     num_rows = df.shape[0]
    #     print(f"Column check {col_check} = {num_rows}")
    #     col_check +=1
        
    df = df[df['PROCESS_REV'] == rev]

    # if debug:
    #     num_rows = df.shape[0]
    #     print(f"Column check {col_check} = {num_rows}")
    #     col_check +=1

    df['SORT_DATE'] = pd.to_datetime(df['SORT_DATE'], errors='coerce')
    # Calculate the start and end dates for the last 4 full weeks
    today = datetime.today()
    start_of_this_week = today - timedelta(days=today.weekday() + 1)
    start_of_4th_last_full_week = start_of_this_week - timedelta(weeks=4)

    # Filter the DataFrame for the last 4 full weeks
    df = df[
        (df['SORT_DATE'] >= start_of_4th_last_full_week)
    ]
    # if debug:
    #     num_rows = df.shape[0]
    #     print(f"Column check {col_check} = {num_rows}")
    #     col_check +=1

    possible_columns = ['IDV', 'SICC', 'CAPABILITY', 'CDYN']
    columns_of_interest = []
    
    for col in possible_columns:        
        if pd.notna(details[col]) and details[col] != '':
            columns_of_interest.append(col)
            df[col] = df[details[col]]
    
    means = df.groupby('FAB')[columns_of_interest].mean().reset_index()
    means['TECH'] = tech
    means['PRODUCT'] = product
    # print(means)

    # for col in columns_of_interest:
    #     if col == 'SICC':
    #         mult = 1000
    #         if sicc_unit != 'mA': mult=1
    #         means['SICC'] = means['SICC']*mult
    #         digits = 2
    #         # means['SICC'] = means['SICC'].round(2)
    #     elif col == 'CDYN':
    #         digits = 2
    #     else:
    #         if idv_unit == 'Mhz':
    #             digits = 0
    #         else:
    #             digits = 2
    #     means[col] = means[col].round(digits)
            

 

        
    
    # for col in columns_of_interest:
        # means[col+'_Flag'] = 'Green'
        
    for col in columns_of_interest:
    # print(col)
        if debug: print(f'  Parameter={col}')
        limit_type = limits[col+'_TYPE']
        target = details[col+'_TGT']
        r_limit = target*(1+limits[col+'_RED'])
        y_limit = target*(1+limits[col+'_YELLOW'])
        
        means[col+'_FLAG'] = means[col].apply(set_flags, args=(limit_type, r_limit, y_limit, debug))
    
    if sicc_unit=='mA':
        means['SICC'] = means['SICC']*1000
    
    # if idv_unit == 'NONE':
    #     means['IDV'] = means['IDV'].round(2)
    #     means['CAPABILITY'] = means['CAPABILITY'].round(2)
    # else:
    #     means['IDV'] = means['IDV'].astype(int)
    #     means['CAPABILITY'] = means['CAPABILITY'].astype(int)



    return means
    
    
def load_excel_to_dict(file_path,key_col):
    df = pd.read_csv(file_path)
    temp_dict = {}
    for _, row in df.iterrows():
        key = row[key_col]
        temp_dict[key] = row.drop(key_col).to_dict()
    return temp_dict
# class Opeda_Scraper:
    
#     def __init__(self, url):
#         self.url = url
#         # if not os.path.exists(download_dir):
#         #     os.makedirs(download_dir)
#         # self.download_dir = download_dir
#         # self.temp_dir = temp_dir
        
#         # self.clear_download_dir()
#         self.options = webdriver.ChromeOptions()
#         self.options.add_argument("--start-maximized")
#         self.options.add_experimental_option("prefs", {
#             # "download.default_directory": self.download_dir,
#             "download.prompt_for_download": False,
#             "download.directory_upgrade": True,
#             "safebrowsing.enabled": True
#         })
#         self.options.add_argument('--no-proxy-server')
#         self.driver = webdriver.Chrome(options=self.options)
#         self.driver.implicitly_wait(10)
#         self.vars = {}

#     def clear_download_dir(self):
#         for filename in os.listdir(self.download_dir):
#             file_path = os.path.join(self.download_dir, filename)
#             try:
#                 if os.path.isfile(file_path) or os.path.islink(file_path):
#                     os.unlink(file_path)
#                 elif os.path.isdir(file_path):
#                     shutil.rmtree(file_path)
#             except Exception as e:
#                 print(f'Failed to delete {file_path}. Reason: {e}')
#         os.makedirs(self.download_dir, exist_ok=True)
        
#     def pull_product_data(self, prod):
#         self.driver.get(self.url)
#         time.sleep(5)
        
#         element = self.driver.find_element(By.XPATH, f"//p[text()='{prod}']")

#         parent_element = element.find_element(By.XPATH, "./ancestor::a")
#         parent_element.click()
#         time.sleep(5)

#         html = self.driver.page_source
        
#         soup = BeautifulSoup(html, 'html.parser')
        
#         self.rows = soup.find_all('div', role='row')

#     def getSupplyRatio(self, sku, operation):
#         for row in self.rows:
#             cells = row.find_all('div', role='gridcell')
#             if len(cells) >9:
#                 column_0_title = cells[0].get('title')
#                 column_2_title = cells[2].get('title')
#                 column_9_title = cells[9].get('title')
#                 if column_0_title == sku and column_2_title == operation:
#                     return column_9_title

def main():

    prod_dict = load_excel_to_dict(config_file_path,'PRODUCT')
    limit_dict = load_excel_to_dict(limits_file_path,'TECH')
    # prod_dict
    
    # opeda = Opeda_Scraper("https://opeda.intel.com/binsplit")

    temp_dfs = []
    for product, details in prod_dict.items():
    # if True:
    #     product = 'RPL68'
        # details = prod_dict['RPL68']
        print(product)
        limits = limit_dict[details['TECH']]
        # product = details['PART']

        # opedaProd = details['OPEDA_PROD']
        # opedaOper = details['OPEDA_OP']
        # supplyRatio = ""
        # firstSKU = True

        # if pd.notna(opedaProd):
        #     skus = details['SKUS'].split(',')
        #     opeda.pull_product_data(opedaProd)
        #     for sku in skus:
        #         if firstSKU:
        #             supplyRatio = f"{sku} = {opeda.getSupplyRatio(sku, opedaOper)}"
        #             firstSKU = False
        #         else:
        #             supplyRatio += f"\n{sku} = {opeda.getSupplyRatio(sku, opedaOper)}"

        temp_df = load_prod_data(server=server, product=product, details=details,limits=limits, idv_unit=details['IDV_UNIT'], sicc_unit=details['SICC_UNIT'], debug=False)

        # temp_df['SupplyRatio'] = supplyRatio

        temp_dfs.append(temp_df)
        
    final_df = pd.concat(temp_dfs, ignore_index=True)

    desired_order = ['TECH','PRODUCT','FAB','IDV','SICC','CAPABILITY','CDYN','IDV_FLAG','SICC_FLAG','CAPABILITY_FLAG','CDYN_FLAG']

    for col in ['IDV','SICC','CAPABILITY','CDYN']:
        if col in final_df.columns:
            final_df[col] = final_df[col].round(2)

    final_df = final_df[desired_order]
    # final_df
    final_df.to_csv(r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\output.csv")



    pythoncom.CoInitialize()  # Initialize COM for the current thread
    excel = win32com.client.Dispatch("Excel.Application")
    filePath = r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\Customer_Ops_Scorecard.xlsm"
    workbook = excel.Workbooks.Open(filePath)
    excel.Application.Run("Convert")
    excel.Quit()
    pythoncom.CoUninitialize()

    os.remove(r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\output.csv")



if __name__ == "__main__":
    main()


    # You can add more functionality or tests here if needed.