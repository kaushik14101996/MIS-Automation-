import pandas as pd
import json
import requests
from xlsxwriter import Workbook
import openpyxl
import streamlit as st
from io import BytesIO
from datetime import timedelta
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import base64
import warnings
warnings.filterwarnings("ignore")
 
# Login Function
# Import All the Necessary Library
 
 
# Login Function
def login():
    st.sidebar.title("Enter the username and password")
    username = st.sidebar.text_input("Enter your username")
    password = st.sidebar.text_input("Enter your password", type="password")
    login_button = st.sidebar.button("Login")
    return username, password, login_button
 
correct_username = "a"
correct_password = "a"
 
# Username and password Check Function
def login_check(username, password):
    return username == correct_username and password == correct_password
 
# Download Function

def convert_df(c1,l1,B2B_Report,B2C_Report,Report,selected_columns_df1):
    with pd.ExcelWriter('selected.xlsx', engine='openpyxl') as writer:
        c1.to_excel(writer, index=False, sheet_name='CC Dump')
        l1.to_excel(writer, index=False, sheet_name='Master')
        if selected_columns_df1:
            B2B_Report[selected_columns_df1].to_excel(writer, index=False, sheet_name='B2B')
        else:
            B2B_Report.to_excel(writer, index=False, sheet_name='B2B')
       
        if selected_columns_df1:
            B2C_Report[selected_columns_df1].to_excel(writer, index=False, sheet_name='B2C')
        else:
            B2C_Report.to_excel(writer, index=False, sheet_name='B2C')
           
        if selected_columns_df1:
            Report[selected_columns_df1].to_excel(writer, index=False, sheet_name='Total')
        else:
            Report.to_excel(writer, index=False, sheet_name='Total')    
 
# Processing function ----------------------------------------------------------------  
@st.cache_data
def preprocess_data(CC_dump_upload,Logistics_Dimension_upload):
    CC_dump = pd.read_excel(CC_dump_upload, dtype='object',parse_dates=['Posting Date'])
    Logistics_Dimension = pd.read_excel(Logistics_Dimension_upload, dtype='object')
    c1=pd.DataFrame(CC_dump)
    l1=pd.DataFrame(Logistics_Dimension)
    Logistics_Dimension = Logistics_Dimension.drop_duplicates()
    Logistics_Dimension['Concat'] = Logistics_Dimension['Company Code'].astype('str') + Logistics_Dimension['Cost Center'] + Logistics_Dimension['GL Code'].astype('str')
    Logistics_Dimension['r'] = Logistics_Dimension.groupby(by = 'Concat')['MIS Classification'].rank(ascending = True, method = 'min')
    Logistics_Dimension = Logistics_Dimension[Logistics_Dimension['r'] == 1]
    Logistics_Dimension = Logistics_Dimension.rename(columns = {'GL Code' : 'Cost Element'})
    Logistics_Dimension['MIS Classification'].fillna('Default Value', inplace=True)
    Logistics_Dimension['Company Code']  = Logistics_Dimension['Company Code'].astype('str')
    Logistics_Dimension['Cost Center']  = Logistics_Dimension['Cost Center'].astype('str')
    Logistics_Dimension['Cost Element']  = Logistics_Dimension['Cost Element'].astype('str')
    Logistics_Dimension['Company Code'] = Logistics_Dimension['Company Code'].str.strip()
    Logistics_Dimension['Cost Center'] = Logistics_Dimension['Cost Center'].str.strip()
    Logistics_Dimension['Cost Element'] = Logistics_Dimension['Cost Element'].str.strip()
    CC_dump['Cost Element'] = CC_dump['Cost Element'].astype('str')
    CC_dump['Cost Center'] = CC_dump['Cost Center'].astype('str')
    CC_dump['Company Code'] = CC_dump['Company Code'].astype('str')
    CC_dump['Company Code'] = CC_dump['Company Code'].str.strip()
    CC_dump['Cost Center'] = CC_dump['Cost Center'].str.strip()
    CC_dump['Cost Element'] = CC_dump['Cost Element'].str.strip()
    merged_data = pd.merge(CC_dump, Logistics_Dimension, on = ['Company Code', 'Cost Center', 'Cost Element'], how = 'left')
    merged_data['Concat'] = merged_data['Company Code'].astype('str') + merged_data['Cost Center'] + merged_data['Cost Element'].astype('str')
    merged_data['Posting Date'] = pd.to_datetime(merged_data['Posting Date'], format='%d-%m-%Y', errors='coerce')
    merged_data_1=merged_data[merged_data['MIS Classification'].isna()][['Concat']]
    merged_data_1=merged_data_1.reset_index(drop=True)
    merged_data_1 = merged_data_1.drop_duplicates()
    if len(merged_data[merged_data['MIS Classification'].isna()]) != 0:
        merged_data = merged_data[merged_data['MIS Classification'].notna()]  
    merged_data['month_year'] = merged_data['Posting Date'].dt.strftime('%b-%y')
    merged_data['year'] = merged_data['Posting Date'].dt.strftime('%y')
    merged_data['month'] = merged_data['Posting Date'].dt.month
    merged_data.loc[merged_data['MIS Classification'].str.contains('B2B'), 'Type'] = 'B2B'
    merged_data.loc[merged_data['MIS Classification'].str.contains('B2C'), 'Type'] = 'B2C' 
    # New Code Changes
    
    merged_data.loc[merged_data["Document Header Text"].isnull(), "Document Header Text"] = ""
    merged_data["Document Header Text"] = merged_data["Document Header Text"].str.lower()
    date_threshold = pd.to_datetime("2024-02-01")
    merged_data.loc[((merged_data["Document Header Text"].str.contains("warehouse charges") &
                 merged_data["MIS Classification"].str.contains("B2C")) &
                 (merged_data["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2C"
    merged_data.loc[((merged_data["Document Header Text"].str.contains("warehouse charges") &
                 merged_data["MIS Classification"].str.contains("B2B")) &
                 (merged_data["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2B"
    merged_data.loc[((merged_data["Document Header Text"].str.contains("wms prov") &
                     merged_data["MIS Classification"].str.contains("B2C")) &
                     (merged_data["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2C"
    merged_data.loc[((merged_data["Document Header Text"].str.contains("wms prov") &
                     merged_data["MIS Classification"].str.contains("B2B")) &
                     (merged_data["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2B"
    
    # Changes Closed
    B2C_merged_data_new = merged_data[merged_data['Type'] == 'B2C']
    B2B_merged_data_new = merged_data[merged_data['Type'] == 'B2B']
    merged_data = pd.DataFrame(merged_data)
    B2C_merged_data_new = pd.DataFrame(B2C_merged_data_new)
    B2B_merged_data_new = pd.DataFrame(B2B_merged_data_new)
    return merged_data,B2C_merged_data_new,B2B_merged_data_new,c1,l1,CC_dump, Logistics_Dimension,merged_data_1
 
# B2B Functions --------------------------------
@st.cache_data
def fun_B2B(merged_data,B2B_merged_data_new):
    if len(B2B_merged_data_new) != 0:
        B2B_merged_data_new = merged_data[merged_data['Type'] == 'B2B']
        B2B_merged_data_new.loc[(B2B_merged_data_new['Document Header Text'] == 'reclass to inward freight') &
                    (B2B_merged_data_new['Concat'] == '1380C13801000466010039'), 'MIS Classification'] = 'Logistics and freight - B2B - Inward'
        B2B_merge_data_CN = B2B_merged_data_new[(B2B_merged_data_new['MIS Classification'] == 'Logistics and freight - B2B')&
                                                (B2B_merged_data_new['Document Header Text'].str.contains('cn'))].copy()
        B2B_merge_data_CN_Freight = B2B_merged_data_new[(B2B_merged_data_new['MIS Classification'] == 'Logistics and freight - B2B')&
                                                        (B2B_merged_data_new['Document Header Text'].str.contains('cn'))&
                                                        (B2B_merged_data_new['Document Header Text'].str.contains('freight'))].copy()
        B2B_merge_data_CN_Freight['MIS Classification'] = 'CN - Lost/Discount'
        B2B_merged_data_new.loc[(B2B_merged_data_new['MIS Classification'] == 'Logistics and freight - B2B') &
                                (B2B_merged_data_new['Document Header Text'].str.contains('cn')), 'MIS Classification'] = 'Logistics & freight - Credit notes'
        B2B_merged_data_new = pd.concat([B2B_merged_data_new,B2B_merge_data_CN,B2B_merge_data_CN_Freight], axis = 0)
        # making Quarters
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 1) &(B2B_merged_data_new['month'] <= 3), 'Quarter'] = 'Q1'
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 4) &(B2B_merged_data_new['month'] <= 6), 'Quarter'] = 'Q2'
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 7) &(B2B_merged_data_new['month'] <= 9), 'Quarter'] = 'Q3'
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 10) &(B2B_merged_data_new['month'] <= 12), 'Quarter'] = 'Q4'
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 1) &(B2B_merged_data_new['month'] <= 6), 'Half yearly'] = 'HY1'
        B2B_merged_data_new.loc[(B2B_merged_data_new['month'] >= 7) &(B2B_merged_data_new['month'] <= 12), 'Half yearly'] = 'HY2'
 
        B2B_merged_data_new['Quarter'] = B2B_merged_data_new['Quarter'] + '-' + B2B_merged_data_new['year']
        B2B_merged_data_new['Half yearly'] = B2B_merged_data_new['Half yearly'] + '-' + B2B_merged_data_new['year']
 
        B2B_PT_data = pd.pivot_table(B2B_merged_data_new, values = 'Value TranCurr', index = ['Type','MIS Classification'],
                   columns = ['month_year'], aggfunc = 'sum', margins = True,margins_name = 'Net Total', fill_value = 0).reset_index()
        B2B_PT_data = B2B_PT_data.drop(['Net Total'], axis = 1)
        B2B_PT_data_Q = B2B_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                           columns = ['Quarter'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
        B2B_PT_data_Q = B2B_PT_data_Q.drop(['Net Total'], axis = 1)
        B2B_PT_data_HY = B2B_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                           columns = ['Half yearly'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
        B2B_PT_data_HY = B2B_PT_data_HY.drop(['Net Total'], axis = 1)
        B2B_PT_data_Y = B2B_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                               columns = ['year'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
        B2B_PT_data_Y = B2B_PT_data_Y.drop(['Net Total'], axis = 1)
        months = [x for x in merged_data['month_year'].unique()]
        for a in months:
            B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
            B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
        Quarter = [x for x in B2B_merged_data_new['Quarter'].unique()]
        for a in Quarter:
            B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
            B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
        half_yearly = [x for x in B2B_merged_data_new['Half yearly'].unique()]
        for a in half_yearly:
            B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
            B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
        yearly = [x for x in B2B_merged_data_new['year'].unique()]
        for a in yearly:
            B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
            B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
        B2B_Report = pd.concat([B2B_PT_data, B2B_PT_data_HY.iloc[:,2:], B2B_PT_data_Q.iloc[:,2:], B2B_PT_data_Y.iloc[:,2:]], axis = 1)
        # New
        year = [x for x in B2B_merged_data_new['year'].unique()]
        year.sort()
        new = []
        pattern = ['Jan-','Feb-', 'Mar-', 'Q1-','Apr-', 'May-', 'Jun-', 'Q2-','HY1-','Jul-', 'Aug-', 'Sep-', 'Q3-','Oct-', 'Nov-', 'Dec-','Q4-','HY2-','']
        col = [k for k in B2B_Report.iloc[:,2:].columns]
        for i in year:
            for j in pattern:
                if (j+i) in col:
                    new.append(j+i)
        B2B_Report = pd.concat([B2B_Report.iloc[:,:2],B2B_Report[new]], axis = 1)
        for i in new:
            B2B_Report.loc[B2B_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i] = -1*B2B_Report.loc[B2B_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i].iloc[0]
            B2B_Report.loc[B2B_Report['Type'] == 'Net Total',i] = B2B_Report.loc[B2B_Report['Type'] == 'Net Total',i].iloc[0] + 2*(B2B_Report.loc[B2B_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i].iloc[0]) + B2B_Report.loc[B2B_Report['MIS Classification'] == 'Logistics and freight - B2B - Inward',i].iloc[0]
        CN_DN = B2B_Report[B2B_Report['MIS Classification'] == 'CN - Lost/Discount']
        Net_Total = B2B_Report[B2B_Report['Type'] == 'Net Total']
        B2B_Report = B2B_Report.drop(B2B_Report[B2B_Report['Type'] == 'Net Total'].index, axis = 0).reset_index(drop = True)
        B2B_Report = B2B_Report.drop(B2B_Report[B2B_Report['MIS Classification'] == 'CN - Lost/Discount'].index, axis = 0).reset_index(drop = True)
        l = len(B2B_Report)
        B2B_Report.loc[l,'Type'] = 'Grand Total'
        B2B_Report.loc[l,'MIS Classification'] = ''
        for i in new:
            B2B_Report.loc[l,i] = B2B_Report[i][0:l].sum()
        B2B_Report = pd.concat([B2B_Report, CN_DN, Net_Total], axis =0).reset_index(drop = True)
        B2B_Report = pd.DataFrame(B2B_Report)
        columns_to_convert = B2B_Report.columns[2:]
        B2B_Report[columns_to_convert] = B2B_Report[columns_to_convert].astype(float)
        numeric_columns = B2B_Report.select_dtypes(include=['number']).columns
        B2B_Report[numeric_columns] = B2B_Report[numeric_columns].applymap(lambda x: round(x / 10000000, 2))
        def custom_sort_key(x):
            if x == 'Logistics and freight - B2B - Inward':
                return 0  
            elif x == 'Logistics and freight - B2B':
                return 1  
            elif x == 'Logistics & freight - Credit notes':
                return 2
            elif x == 'WH Rent - B2B':
                return 3
            elif x == 'Insurance - B2B':
                return 4
            elif x == 'Salary - B2B':
                return 5
            elif x == 'Travel & Others - B2B':
                return 6
            elif x == 'B2B - ESOP':
                return 7
            else:
                return 8
        B2B_Report['sort_key'] = B2B_Report['MIS Classification'].apply(custom_sort_key)  
        B2B_Report = B2B_Report.sort_values(by='sort_key')
        B2B_Report = B2B_Report.drop(columns=['sort_key'])
          
            
    return B2B_Report
           
           
   
   
   
   
   
 
# B2C Functions --------------------------------
@st.cache_data
def fun_B2C(merged_data,B2C_merged_data_new):
    if len(B2C_merged_data_new) != 0:
        B2C_merged_data_new = merged_data[merged_data['Type'] == 'B2C']
        B2C_merged_data_new.loc[(B2C_merged_data_new['Document Header Text'] == 'reclass to inward freight') &
                    (B2C_merged_data_new['Concat'] == '1380C13801000466010039'), 'MIS Classification'] = 'Logistics and freight - B2C - Inward'
        B2C_merged_data_new_CN = B2C_merged_data_new[(B2C_merged_data_new['MIS Classification'] == 'Logistics and freight - B2C')
                                            &(B2C_merged_data_new['Document Header Text'].str.contains('cn'))].copy()
        B2C_merged_data_new_CN_LD = B2C_merged_data_new[(B2C_merged_data_new['MIS Classification'] == 'Logistics and freight - B2C')
                                            &(B2C_merged_data_new['Document Header Text'].str.contains('cn'))
                                            &(B2C_merged_data_new['Document Header Text'].str.contains('freight'))].copy()
        B2C_merged_data_new_CN_LD['MIS Classification'] = 'Credit note - for discount'
        B2C_merged_data_new.loc[(B2C_merged_data_new['MIS Classification'] == 'Logistics and freight - B2C') &
                    (B2C_merged_data_new['Document Header Text'].str.contains('cn')), 'MIS Classification'] = 'Logistics & freight - Credit notes'
        B2C_merged_data_new = pd.concat([B2C_merged_data_new, B2C_merged_data_new_CN,B2C_merged_data_new_CN_LD], axis = 0)
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 1) &(B2C_merged_data_new['month'] <= 3), 'Quarter'] = 'Q1'
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 4) &(B2C_merged_data_new['month'] <= 6), 'Quarter'] = 'Q2'
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 7) &(B2C_merged_data_new['month'] <= 9), 'Quarter'] = 'Q3'
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 10) &(B2C_merged_data_new['month'] <= 12), 'Quarter'] = 'Q4'
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 1) &(B2C_merged_data_new['month'] <= 6), 'Half yearly'] = 'HY1'
        B2C_merged_data_new.loc[(B2C_merged_data_new['month'] >= 7) &(B2C_merged_data_new['month'] <= 12), 'Half yearly'] = 'HY2'
 
        B2C_merged_data_new['Quarter'] = B2C_merged_data_new['Quarter'] + '-' + B2C_merged_data_new['year']
        B2C_merged_data_new['Half yearly'] = B2C_merged_data_new['Half yearly'] + '-' + B2C_merged_data_new['year']
 
        B2C_PT_data = pd.pivot_table(B2C_merged_data_new, values = 'Value TranCurr', index = ['Type','MIS Classification'],
                   columns = ['month_year'], aggfunc = 'sum', margins = True,margins_name = 'Net Total',fill_value = 0).reset_index()
        B2C_PT_data = B2C_PT_data.drop(['Net Total'], axis = 1)
        B2C_PT_data_Q = B2C_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                           columns = ['Quarter'], aggfunc = 'sum', margins = True, margins_name = 'Net Total',fill_value = 0).reset_index()
        B2C_PT_data_Q = B2C_PT_data_Q.drop(['Net Total'], axis = 1)
        B2C_PT_data_HY = B2C_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                           columns = ['Half yearly'], aggfunc = 'sum', margins = True, margins_name = 'Net Total',fill_value = 0).reset_index()
        B2C_PT_data_HY = B2C_PT_data_HY.drop(['Net Total'], axis = 1)
        B2C_PT_data_Y = B2C_merged_data_new.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                               columns = ['year'], aggfunc = 'sum', margins = True, margins_name = 'Net Total',fill_value = 0).reset_index()
        B2C_PT_data_Y = B2C_PT_data_Y.drop(['Net Total'], axis = 1)
        B2C_Report = pd.concat([B2C_PT_data, B2C_PT_data_HY.iloc[:,2:], B2C_PT_data_Q.iloc[:,2:], B2C_PT_data_Y.iloc[:,2:]], axis = 1)
        
        
        
        year = [x for x in B2C_merged_data_new['year'].unique()]
        year.sort()
        new = []
        pattern = ['Jan-','Feb-', 'Mar-', 'Q1-','Apr-', 'May-', 'Jun-', 'Q2-','HY1-','Jul-', 'Aug-', 'Sep-', 'Q3-','Oct-', 'Nov-', 'Dec-','Q4-','HY2-','']
        col = [k for k in B2C_Report.iloc[:,2:].columns]
        for i in year:
            for j in pattern:
                if (j+i) in col:
                    new.append(j+i)
                    
        for i in new:
            B2C_Report.loc[B2C_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i] = -1*B2C_Report.loc[B2C_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i].iloc[0]
            B2C_Report.loc[B2C_Report['Type'] == 'Net Total',i] = B2C_Report.loc[B2C_Report['Type'] == 'Net Total',i].iloc[0] + 2*(B2C_Report.loc[B2C_Report['MIS Classification'] == 'Logistics & freight - Credit notes',i].iloc[0])
        CN_DN = B2C_Report[B2C_Report['MIS Classification'] == 'Credit note - for discount']
        
        Net_Total = B2C_Report[B2C_Report['Type'] == 'Net Total']
        B2C_Report = B2C_Report.drop(B2C_Report[B2C_Report['Type'] == 'Net Total'].index, axis = 0).reset_index(drop = True)
        B2C_Report = B2C_Report.drop(B2C_Report[B2C_Report['MIS Classification'] == 'Credit note - for discount'].index, axis = 0).reset_index(drop = True)
        
        
        l = len(B2C_Report)
        B2C_Report.loc[l,'Type'] = 'Grand Total'
        B2C_Report.loc[l,'MIS Classification'] = ''
        for i in new:
            B2C_Report.loc[l,i] = B2C_Report[i][0:l].sum()
        B2C_Report = pd.concat([B2C_Report, CN_DN, Net_Total], axis = 0).reset_index(drop = True)
                
        
        B2C_Report = pd.DataFrame(B2C_Report)
        columns_to_convert = B2C_Report.columns[2:]
        B2C_Report[columns_to_convert] = B2C_Report[columns_to_convert].astype(float)
        numeric_columns = B2C_Report.select_dtypes(include=['number']).columns
        B2C_Report[numeric_columns] = B2C_Report[numeric_columns].applymap(lambda x: round(x / 10000000, 2))   
        
        def custom_sort_key(x):
            if x == 'Logistics and freight - B2C':
                return 0  
            elif x == 'Logistics & freight - Credit notes':
                return 1  
            elif x == 'WH Rent - B2C':
                return 2
            elif x == 'Packaging - B2C':
                return 3
            elif x == 'Salary - B2C':
                return 4
            elif x == 'Travel & Others - B2C':
                return 5
            else:
                return 6
        B2C_Report['sort_key'] = B2C_Report['MIS Classification'].apply(custom_sort_key)  
        B2C_Report = B2C_Report.sort_values(by='sort_key')
        B2C_Report = B2C_Report.drop(columns=['sort_key'])
         
    return  B2C_Report       
   
   
   
   
   
   
# Add function Start    
@st.cache_data
def MIS_add(CC_dump,Logistics_Dimension):
    merged_data_new = pd.merge(CC_dump, Logistics_Dimension, on = ['Company Code', 'Cost Center', 'Cost Element'], how = 'left')
    # New
    
    merged_data_new.loc[merged_data_new["Document Header Text"].isnull(), "Document Header Text"] = ""
    merged_data_new.loc[:, "Document Header Text"] = merged_data_new["Document Header Text"].str.lower()
    date_threshold = pd.to_datetime("2024-02-01")
    merged_data_new.loc[((merged_data_new["Document Header Text"].str.contains("warehouse charges") &
                 merged_data_new["MIS Classification"].str.contains("B2C")) &
                 (merged_data_new["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2C"
    merged_data_new.loc[((merged_data_new["Document Header Text"].str.contains("warehouse charges") &
                 merged_data_new["MIS Classification"].str.contains("B2B")) &
                 (merged_data_new["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2B"
    merged_data_new.loc[((merged_data_new["Document Header Text"].str.contains("wms prov") &
                     merged_data_new["MIS Classification"].str.contains("B2C")) &
                     (merged_data_new["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2C"
    merged_data_new.loc[((merged_data_new["Document Header Text"].str.contains("wms prov")) &
                     (merged_data_new["MIS Classification"].str.contains("B2B")) &
                     (merged_data_new["Posting Date"] >= date_threshold)), "MIS Classification"] = "WH Rent - B2B"
    
    
    
    
    # Close
    
    merged_data_new.loc[(merged_data_new['Document Header Text'] == 'reclass to inward freight') &
                (merged_data_new['Concat'] == '1380C13801000466010039'), 'MIS Classification'] = 'Logistics and freight - B2B - Inward'
    merged_data_new = merged_data_new[merged_data_new['MIS Classification'] != 'Logistics and freight - B2B - Inward']
    credit_note = merged_data_new[(merged_data_new['MIS Classification'].str.contains('Logistics and freight'))
                          &(merged_data_new['Document Header Text'].str.contains('cn'))].copy()
    credit_note['Value TranCurr'] = -credit_note['Value TranCurr']
    CN_DN = merged_data_new[(merged_data_new['MIS Classification'].str.contains('Logistics and freight')) &
                    (merged_data_new['Document Header Text'].str.contains('cn'))
                    &(merged_data_new['Document Header Text'].str.contains('freight'))].copy()
    CN_DN['MIS Classification'] = 'CN - Lost/Discount'
    merged_data_new = pd.concat([merged_data_new, credit_note], axis = 0)
    merged_data_new['Posting Date'] = pd.to_datetime(merged_data_new['Posting Date'], format='%d-%m-%Y', errors='coerce')
    # merged_data_new['Posting Date'] =pd.to_datetime(merged_data_new['Posting Date'])
    merged_data_new['month_year'] = merged_data_new['Posting Date'].dt.strftime('%b-%y')
    merged_data_new['year'] = merged_data_new['Posting Date'].dt.strftime('%y')
    merged_data_new['month'] = merged_data_new['Posting Date'].dt.month
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('B2B'), 'Type'] = 'B2B'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('B2C'), 'Type'] = 'B2C'


    CN_DN = CN_DN.reset_index()
    
    if len(CN_DN) > 0:   
        CN_DN['Posting Date'] = pd.to_datetime(CN_DN['Posting Date'], format='%d-%m-%Y', errors='coerce')
        CN_DN['month_year'] = CN_DN['Posting Date'].dt.strftime('%b-%y')
        CN_DN['year'] = CN_DN['Posting Date'].dt.strftime('%y')
        CN_DN['month'] = CN_DN['Posting Date'].dt.month
        CN_DN.loc[CN_DN['MIS Classification'].str.contains('B2B'), 'Type'] = 'B2B'
        CN_DN.loc[CN_DN['MIS Classification'].str.contains('B2C'), 'Type'] = 'B2C'
        
    
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('Salary'), 'MIS Classification'] = 'Salary'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('Travel & Others'), 'MIS Classification'] = 'Travel & Others'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('WH Rent'), 'MIS Classification'] = 'WH Rent'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('Packaging'), 'MIS Classification'] = 'Packaging'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('Logistics and freight'), 'MIS Classification'] = 'Logistics and freight'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('Insurance'), 'MIS Classification'] = 'Insurance'
    merged_data_new.loc[merged_data_new['MIS Classification'].str.contains('ESOP'), 'MIS Classification'] = 'ESOP'
    merged_data_new = pd.concat([merged_data_new, CN_DN], axis = 0)
    
    
    merged_data_new.loc[(merged_data_new['month'] >= 1) &(merged_data_new['month'] <= 3), 'Quarter'] = 'Q1'
    merged_data_new.loc[(merged_data_new['month'] >= 4) &(merged_data_new['month'] <= 6), 'Quarter'] = 'Q2'
    merged_data_new.loc[(merged_data_new['month'] >= 7) &(merged_data_new['month'] <= 9), 'Quarter'] = 'Q3'
    merged_data_new.loc[(merged_data_new['month'] >= 10) &(merged_data_new['month'] <= 12), 'Quarter'] = 'Q4'
    merged_data_new.loc[(merged_data_new['month'] >= 1) &(merged_data_new['month'] <= 6), 'Half yearly'] = 'HY1'
    merged_data_new.loc[(merged_data_new['month'] >= 7) &(merged_data_new['month'] <= 12), 'Half yearly'] = 'HY2'
 
    merged_data_new['Quarter'] = merged_data_new['Quarter'] + '-' + merged_data_new['year']
    merged_data_new['Half yearly'] = merged_data_new['Half yearly'] + '-' + merged_data_new['year']
 
    PT_data = pd.pivot_table(merged_data_new, values = 'Value TranCurr', index = ['MIS Classification'],
               columns = ['month_year'], aggfunc = 'sum', margins = True,margins_name = 'Net Total', fill_value = 0).reset_index()
    PT_data = PT_data.drop(['Net Total'], axis = 1)
    PT_data_Q = merged_data_new.pivot_table(values = 'Value TranCurr', index = [ 'MIS Classification'],
                       columns = ['Quarter'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
    PT_data_Q = PT_data_Q.drop(['Net Total'], axis = 1)
    PT_data_HY = merged_data_new.pivot_table(values = 'Value TranCurr', index = [ 'MIS Classification'],
                       columns = ['Half yearly'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
    PT_data_HY = PT_data_HY.drop(['Net Total'], axis = 1)
    PT_data_Y = merged_data_new.pivot_table(values = 'Value TranCurr', index = [ 'MIS Classification'],
                                           columns = ['year'], aggfunc = 'sum', margins = True, margins_name = 'Net Total', fill_value = 0).reset_index()
    PT_data_Y = PT_data_Y.drop(['Net Total'], axis = 1)
    Report = pd.concat([PT_data, PT_data_HY.iloc[:,1:], PT_data_Q.iloc[:,1:], PT_data_Y.iloc[:,1:]], axis = 1)
    year = [x for x in merged_data_new['year'].unique()]
    year.sort()
    new = []
    pattern = ['Jan-','Feb-', 'Mar-', 'Q1-','Apr-', 'May-', 'Jun-', 'Q2-','HY1-','Jul-', 'Aug-', 'Sep-', 'Q3-','Oct-', 'Nov-', 'Dec-','Q4-','HY2-','']
    col = [k for k in Report.iloc[:,1:].columns]
    for i in year:
        for j in pattern:
            if (j+i) in col:
                new.append(j+i)
    Report = pd.concat([Report.iloc[:,:1],Report[new]], axis = 1)
 
    CN = Report[Report['MIS Classification'] == 'CN - Lost/Discount']
    Net_Total = Report[Report['MIS Classification'] == 'Net Total']
    Report = Report.drop(Report[Report['MIS Classification'] == 'CN - Lost/Discount'].index, axis = 0).reset_index(drop = True)
    Report = Report.drop(Report[Report['MIS Classification'] == 'Net Total'].index, axis = 0).reset_index(drop = True)
    
    Report_len = len(Report)
    Report.loc[Report_len, 'MIS Classification'] = "Grand Total"
    for j in new:
        Report.loc[Report_len, j] = Report[j][0:Report_len].sum()
 
    numeric_columns = Report.select_dtypes(include=['number']).columns
    Report[numeric_columns] = Report[numeric_columns].applymap(lambda x: '{:.2f}'.format(x / 10000000))
    for i in numeric_columns:
        Report[i] = Report[i].astype('float')
    Report = pd.concat([Report, CN, Net_Total], axis = 0).reset_index(drop = True)
    Report = pd.DataFrame(Report)
    columns_to_convert = Report.columns[1:]
    Report[columns_to_convert] = Report[columns_to_convert].astype(float)
    numeric_columns = Report.select_dtypes(include=['number']).columns
    Report[numeric_columns] = Report[numeric_columns].applymap(lambda x: round(x / 10000000, 2))
    
    def custom_sort_key(x):
      if x == 'Logistics and freight':
          return 0  
      elif x == 'WH Rent':
          return 1  
      elif x == 'Packaging':
          return 2
      elif x == 'Insurance':
          return 3
      elif x == 'Salary':
          return 4
      elif x == 'Travel & Others':
          return 5
      else:
          return 6
    Report['sort_key'] = Report['MIS Classification'].apply(custom_sort_key)  
    Report = Report.sort_values(by='sort_key')
    Report = Report.drop(columns=['sort_key'])    

    return Report
   
# Total Finish
 
 
 
   
 
# Main function Setup
def main():
    #global CC_dump,Logistics_Dimension,B2B_Report,B2C_Report,c1,l1,B2B,B2C,selected_columns_df1,selected_columns_df2
    st.set_page_config(
        page_title="MIS Summary Automation",
        layout="wide",
        page_icon="🧊",
    )
 
    st.markdown("""
    <script>
      document.addEventListener('hideSidebar', function() {
        document.querySelector('.sidebar').style.display = 'none';
      });
    </script>
  """, unsafe_allow_html=True)
 
    st.markdown('<h2 style="text-align: center; font-size: 45px; font-weight: bold;">MIS Summary Automation</h2>', unsafe_allow_html=True)
    st.title("")
 
    username, password, login_button = login()
 
    # Login Check Function
   
    if login_check(username, password):
        st.sidebar.success("Login Successfully")
        st.sidebar.markdown(
            """
            ## Contact Information
            If you encounter any difficulties, please contact:
            - **Name:** Rohit Kaushik, Abhishek Pal
            - **Phone Number:** +91-9654741555
            - **Email:** rohit.kaushik@quation.in,abhishek.pal@quatiom.in
            """
        )
        # File Uploader CC Dump
        st.markdown('<h2 style="text-align: center; font-size: 24px; font-weight: bold;">Upload CC Dump File</h2>', unsafe_allow_html=True)
        CC_dump_upload = st.file_uploader(" Upload CC Dump ", type=["xlsx"])
       
        # File Uploader Mater
        st.markdown('<h2 style="text-align: center; font-size: 25px; font-weight: bold;">Upload Logistic File</h2>', unsafe_allow_html=True)
        Logistics_Dimension_upload = st.file_uploader("Upload  Master File", type=["xlsx"])
       
        # Files Import
        try:
            if CC_dump_upload and Logistics_Dimension_upload:
                login_button = st.button("Start Processing")
                if "load_state" not in st.session_state:
                    st.session_state.load_state = False
                if login_button or st.session_state.load_state:
                    st.session_state.load_state = True
                    merged_data,B2C_merged_data_new,B2B_merged_data_new,c1,l1,CC_dump, Logistics_Dimension,merged_data_1 = preprocess_data(CC_dump_upload, Logistics_Dimension_upload)
                    B2B_Report = fun_B2B(merged_data,B2B_merged_data_new)
                    B2C_Report = fun_B2C(merged_data,B2C_merged_data_new)
                    Report = MIS_add(CC_dump, Logistics_Dimension)
                   
 
                    # Preprocessing COmplete
                   
                    st.subheader("Logistic File")
                    AgGrid(l1, height=400, return_mode='both')
                   
                    st.subheader("CC_dump")
                    AgGrid(c1, height=400, return_mode='both')
 
                    st.subheader("B2B Summary")
                    AgGrid(B2B_Report, height=400, return_mode='both')
                   
                    st.subheader("B2C Summary")
                    AgGrid(B2C_Report, height=400, return_mode='both')
                   
                    st.subheader("B2B + B2C")
                    AgGrid(Report, height=400, return_mode='both')
                   
                    st.subheader("Not Matching Records")
                    AgGrid(merged_data_1, height=400, return_mode='both')
                   
                    excel_buffer = BytesIO()
                    merged_data_1.to_excel(excel_buffer, index=False, engine='openpyxl')
                    excel_buffer.seek(0)
                    st.download_button(
                        label='Download Excel File',
                        data=excel_buffer,
                        file_name='sample_data.xlsx',
                        key='download_excel_button'
                    )
 
                    gb = GridOptionsBuilder.from_dataframe(pd.concat([B2B_Report, B2B_Report, Report], keys=['B2B_Report', 'B2B_Report', 'Report']))
                    gb.configure_default_column(enablePivot=True, enableValue=True, enableRowGroup=True)
                    gb.configure_selection(selection_mode="multiple", use_checkbox=False)
                    gb.configure_side_bar()
                   
 
                    selected_columns_df1 = st.multiselect('Select parameters to download', B2B_Report.columns.tolist())
 
                    convert_df(c1,l1,B2B_Report,B2C_Report,Report,selected_columns_df1)
 
                    with open('selected.xlsx', 'rb') as f:
                        bytes = f.read()
                        b64 = base64.b64encode(bytes).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="selected.xlsx">Download data as Excel</a>'
                        st.markdown(href, unsafe_allow_html=True)
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
# Calling Main Function
if __name__ == "__main__":
    main()
