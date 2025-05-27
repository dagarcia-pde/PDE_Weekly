import pandas as pd
import os
from datetime import datetime, timedelta

config_file_path = 'Config.csv'
temp_file_path = 'Temp.csv'


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
prod_dict = load_excel_to_dict('Config.csv','PRODUCT')
limit_dict = load_excel_to_dict('Tech_Limits.csv','TECH')
# prod_dict
server = 'rasinkul-desk'

temp_dfs = []
for product, details in prod_dict.items():
# if True:
#     product = 'RPL68'
    # details = prod_dict['RPL68']
    print(product)
    limits = limit_dict[details['TECH']]
    # product = details['PART']

    temp_df = load_prod_data(server=server, product=product, details=details,limits=limits, idv_unit=details['IDV_UNIT'], sicc_unit=details['SICC_UNIT'], debug=False)

    temp_dfs.append(temp_df)
    
final_df = pd.concat(temp_dfs, ignore_index=True)

desired_order = ['TECH','PRODUCT','FAB','IDV','SICC','CAPABILITY','CDYN','IDV_FLAG','SICC_FLAG','CAPABILITY_FLAG','CDYN_FLAG']

for col in ['IDV','SICC','CAPABILITY','CDYN']:
    if col in final_df.columns:
        final_df[col] = final_df[col].round(2)

final_df = final_df[desired_order]
# final_df
final_df.to_csv(r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\output.csv")


import win32com.client
import pythoncom
pythoncom.CoInitialize()  # Initialize COM for the current thread
excel = win32com.client.Dispatch("Excel.Application")
filePath = r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\Customer_Ops_Scorecard.xlsm"
workbook = excel.Workbooks.Open(filePath)
excel.Application.Run("Convert")
excel.Quit()
pythoncom.CoUninitialize()

os.remove(r"\\azshfs.intel.com\AZAnalysis$\1272_MAODATA\Config\PDE\dagarcia\PDE_Weekly\output.csv")