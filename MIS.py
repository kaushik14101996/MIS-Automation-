# Import All the Necessary Library
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
def convert_df(c1,l1,B2B,B2C,selected_columns_df1,selected_columns_df2):
    with pd.ExcelWriter('selected.xlsx', engine='openpyxl') as writer:
        c1.to_excel(writer, index=False, sheet_name='CC Dump')
        l1.to_excel(writer, index=False, sheet_name='Master')
        if selected_columns_df1:
            B2B[selected_columns_df1].to_excel(writer, index=False, sheet_name='B2B')
        else:
            B2B.to_excel(writer, index=False, sheet_name='B2B')
        
        if selected_columns_df2:
            B2C[selected_columns_df2].to_excel(writer, index=False, sheet_name='B2C')
        else:
            B2C.to_excel(writer, index=False, sheet_name='B2C')

# Main function
def main():
    global CC_dump,Logistics_Dimension,B2B_Report,B2C_Report,c1,l1,B2B,B2C,selected_columns_df1,selected_columns_df2
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
            - **Name:** Rohit Kaushik
            - **Phone Number:** +91-9654741555
            - **Email:** rohit.kaushik@quation.in
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
                login_button = st.checkbox("Start Processing")
                if login_button:
                    CC_dump = pd.read_excel(CC_dump_upload, dtype='object')
                    c1=pd.DataFrame(CC_dump)
                    Logistics_Dimension = pd.read_excel(Logistics_Dimension_upload, dtype='object')
                    l1=pd.DataFrame(Logistics_Dimension)
                    # Import Complete 
                    # Preprocessing Start
                    Logistics_Dimension = Logistics_Dimension.drop_duplicates()
                    Logistics_Dimension['Concat'] = Logistics_Dimension['Company Code'].astype('str') + Logistics_Dimension['Cost Center'] + Logistics_Dimension['GL Code'].astype('str')
                    Logistics_Dimension['r'] = Logistics_Dimension.groupby(by = 'Concat')['MIS Classification'].rank(ascending = True, method = 'min')
                    Logistics_Dimension = Logistics_Dimension[Logistics_Dimension['r'] == 1]
                    Logistics_Dimension = Logistics_Dimension.rename(columns = {'GL Code' : 'Cost Element'})
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
                    merged_data['Concat'] = merged_data['Company Code'].astype('str') + merged_data['Cost Center'].astype('str') + merged_data['Cost Element'].astype('str')
                    merged_data['Posting Date'] = pd.to_datetime(merged_data['Posting Date'])
                    if len(merged_data[merged_data['MIS Classification'].isna()]) != 0:
                        merged_data = merged_data[merged_data['MIS Classification'].notna()]
                    merged_data['month_year'] = merged_data['Posting Date'].dt.strftime('%b-%y')
                    merged_data['year'] = merged_data['Posting Date'].dt.strftime('%y')
                    merged_data['month'] = merged_data['Posting Date'].dt.month
                    merged_data.loc[merged_data['MIS Classification'].str.contains('B2B'), 'Type'] = 'B2B'
                    merged_data.loc[merged_data['MIS Classification'].str.contains('B2C'), 'Type'] = 'B2C'  
                    B2C_merged_data = merged_data[merged_data['Type'] == 'B2C']
                    B2B_merged_data = merged_data[merged_data['Type'] == 'B2B']



                    # Preprocessing COmplete

                    # B2B Start
                    if len(B2B_merged_data) != 0:
                        B2B_merged_data = merged_data[merged_data['Type'] == 'B2B']
                    B2B_merged_data.loc[(B2B_merged_data['Document Header Text'] == 'Reclass to Inward Freight') &
                                (B2B_merged_data['Concat'] == '1380C13801000466010039'), 'MIS Classification'] = 'Logistics and freight - B2B - Inward'
                    B2B_merged_data.loc[(B2B_merged_data['MIS Classification'] == 'Logistics and freight - B2B') &
                                (B2B_merged_data['Document Header Text'].str.contains('CN')), 'MIS Classification'] = 'Logistics & freight - Credit notes'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 1) &(B2B_merged_data['month'] <= 3), 'Quarter'] = 'Q1'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 4) &(B2B_merged_data['month'] <= 6), 'Quarter'] = 'Q2'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 7) &(B2B_merged_data['month'] <= 9), 'Quarter'] = 'Q3'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 10) &(B2B_merged_data['month'] <= 12), 'Quarter'] = 'Q4'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 1) &(B2B_merged_data['month'] <= 6), 'Half yearly'] = 'HY1'
                    B2B_merged_data.loc[(B2B_merged_data['month'] >= 7) &(B2B_merged_data['month'] <= 12), 'Half yearly'] = 'HY2'

                    B2B_merged_data['Quarter'] = B2B_merged_data['Quarter'] + '-' + B2B_merged_data['year']
                    B2B_merged_data['Half yearly'] = B2B_merged_data['Half yearly'] + '-' + B2B_merged_data['year']

                    B2B_PT_data = pd.pivot_table(B2B_merged_data, values = 'Value TranCurr', index = ['Type','MIS Classification'], 
                               columns = ['month_year'], aggfunc = 'sum', margins = True,margins_name = 'Grand Total').reset_index()
                    B2B_PT_data = B2B_PT_data.drop(['Grand Total'], axis = 1)
                    B2B_PT_data_Q = B2B_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                       columns = ['Quarter'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2B_PT_data_Q = B2B_PT_data_Q.drop(['Grand Total'], axis = 1)
                    B2B_PT_data_HY = B2B_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                       columns = ['Half yearly'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2B_PT_data_HY = B2B_PT_data_HY.drop(['Grand Total'], axis = 1)
                    B2B_PT_data_Y = B2B_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                           columns = ['year'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2B_PT_data_Y = B2B_PT_data_Y.drop(['Grand Total'], axis = 1)
                    months = [x for x in merged_data['month_year'].unique()]
                    for a in months:
                        B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                        B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data.loc[B2B_PT_data['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                    Quarter = [x for x in B2B_merged_data['Quarter'].unique()]
                    for a in Quarter:
                        B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                        B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_Q.loc[B2B_PT_data_Q['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                    half_yearly = [x for x in B2B_merged_data['Half yearly'].unique()]
                    for a in half_yearly:
                        B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                        B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_HY.loc[B2B_PT_data_HY['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                    yearly = [x for x in B2C_merged_data['year'].unique()]
                    for a in yearly:
                        B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B',a] = B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B',a].iloc[0] + B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                        B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a] = -B2B_PT_data_Y.loc[B2B_PT_data_Y['MIS Classification'] == 'Logistics and freight - B2B - Inward',a].iloc[0]
                    B2B_Report = pd.concat([B2B_PT_data, B2B_PT_data_HY.iloc[:,2:], B2B_PT_data_Q.iloc[:,2:], B2B_PT_data_Y.iloc[:,2:]], axis = 1)


                    # B2B_Report[numeric_columns] = B2B_Report[numeric_columns].applymap(lambda x: '{:.2f}'.format(x / 10000000))
                    # for i in numeric_columns:
                    #     B2B_Report[i] = B2B_Report[i].astype('float')

                    new = []
                    pattern = ['Jan-','Feb-', 'Mar-', 'Q1-','Apr-', 'May-', 'Jun-', 'Q2-','HY1-','Jul-', 'Aug-', 'Sep-', 'Q3-','Oct-', 'Nov-', 'Dec-','Q4-','HY2-','']
                    year = [x for x in B2B_merged_data['year'].unique()]
                    year.sort()
                    col = [k for k in B2B_Report.iloc[:,2:].columns]
                    for i in year:
                        for j in pattern:
                            if (j+i) in col:
                                new.append(j+i)
                    numeric_columns = B2B_Report.select_dtypes(include=['number']).columns            
                    B2B_Report[numeric_columns] = B2B_Report[numeric_columns] / 10000000
                    B2B_Report[numeric_columns] = B2B_Report[numeric_columns].applymap(lambda x: '{:.1f}'.format(x))            
                    B2B_Report = pd.concat([B2B_Report.iloc[:,:2],B2B_Report[new]], axis = 1) 
                    B2B = pd.DataFrame(B2B_Report)

                    AgGrid(B2B_Report, height=400, return_mode='both')


                    # B2B Complete 

                    # B2C Start

                    if len(B2C_merged_data) != 0:
                        B2C_merged_data = merged_data[merged_data['Type'] == 'B2C']
                    B2C_merged_data.loc[(B2C_merged_data['Document Header Text'] == 'Reclass to Inward Freight') &
                                (B2C_merged_data['Concat'] == '1380C13801000466010039'), 'MIS Classification'] = 'Logistics and freight - B2C - Inward'
                    B2C_merged_data.loc[(B2C_merged_data['MIS Classification'] == 'Logistics and freight - B2C') &
                                (B2C_merged_data['Document Header Text'].str.contains('CN')), 'MIS Classification'] = 'Logistics & freight - Credit notes'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 1) &(B2C_merged_data['month'] <= 3), 'Quarter'] = 'Q1'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 4) &(B2C_merged_data['month'] <= 6), 'Quarter'] = 'Q2'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 7) &(B2C_merged_data['month'] <= 9), 'Quarter'] = 'Q3'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 10) &(B2C_merged_data['month'] <= 12), 'Quarter'] = 'Q4'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 1) &(B2C_merged_data['month'] <= 6), 'Half yearly'] = 'HY1'
                    B2C_merged_data.loc[(B2C_merged_data['month'] >= 7) &(B2C_merged_data['month'] <= 12), 'Half yearly'] = 'HY2'

                    B2C_merged_data['Quarter'] = B2C_merged_data['Quarter'] + '-' + B2C_merged_data['year']
                    B2C_merged_data['Half yearly'] = B2C_merged_data['Half yearly'] + '-' + B2C_merged_data['year']

                    B2C_PT_data = pd.pivot_table(B2C_merged_data, values = 'Value TranCurr', index = ['Type','MIS Classification'], 
                               columns = ['month_year'], aggfunc = 'sum', margins = True,margins_name = 'Grand Total').reset_index()
                    B2C_PT_data = B2C_PT_data.drop(['Grand Total'], axis = 1)
                    B2C_PT_data_Q = B2C_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                       columns = ['Quarter'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2C_PT_data_Q = B2C_PT_data_Q.drop(['Grand Total'], axis = 1)
                    B2C_PT_data_HY = B2C_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                       columns = ['Half yearly'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2C_PT_data_HY = B2C_PT_data_HY.drop(['Grand Total'], axis = 1)
                    B2C_PT_data_Y = B2C_merged_data.pivot_table(values = 'Value TranCurr', index = ['Type', 'MIS Classification'],
                                                           columns = ['year'], aggfunc = 'sum', margins = True, margins_name = 'Grand Total').reset_index()
                    B2C_PT_data_Y = B2C_PT_data_Y.drop(['Grand Total'], axis = 1)
                    B2C_Report = pd.concat([B2C_PT_data, B2C_PT_data_HY.iloc[:,2:], B2C_PT_data_Q.iloc[:,2:], B2C_PT_data_Y.iloc[:,2:]], axis = 1)

                    # B2C_Report[numeric_columns] = B2C_Report[numeric_columns].applymap(lambda x: '{:.2f}'.format(x / 10000000))
                    # for i in numeric_columns:
                    #     B2C_Report[i] = B2C_Report[i].astype('float') 

                    year = [x for x in B2C_merged_data['year'].unique()]
                    year.sort()   
                    new = []
                    pattern = ['Jan-','Feb-', 'Mar-', 'Q1-','Apr-', 'May-', 'Jun-', 'Q2-','HY1-','Jul-', 'Aug-', 'Sep-', 'Q3-','Oct-', 'Nov-', 'Dec-','Q4-','HY2-','']
                    col = [k for k in B2C_Report.iloc[:,2:].columns]
                    for i in year:
                        for j in pattern:
                            if (j+i) in col:
                                new.append(j+i) 

                    B2C_Report = pd.concat([B2C_Report.iloc[:,:2],B2C_Report[new]], axis = 1)
                    B2C = pd.DataFrame(B2B_Report) 

                    AgGrid(B2C_Report, height=400, return_mode='both')




                    # B2C Complete
                    gb = GridOptionsBuilder.from_dataframe(pd.concat([B2B, B2C], keys=['B2B', 'B2B']))
                    gb.configure_default_column(enablePivot=True, enableValue=True, enableRowGroup=True)
                    gb.configure_selection(selection_mode="multiple", use_checkbox=False)
                    gb.configure_side_bar()
                    gridoptions = gb.build()

                    # response = AgGrid(
                    #     pd.concat([B2B, B2B], keys=['B2B', 'B2B']),
                    #     height=200,
                    #     gridOptions=gridoptions,
                    #     enable_enterprise_modules=True,
                    #     update_mode=GridUpdateMode.MODEL_CHANGED,
                    #     data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                    #     fit_columns_on_grid_load=False,
                    #     header_checkbox_selection_filtered_only=True,
                    #     use_checkbox=True)
                    selected_columns_df1 = st.multiselect('Select columns from B2B to download', B2B.columns.tolist())
                    selected_columns_df2 = st.multiselect('Select columns from B2C to download', B2C.columns.tolist())

                    convert_df(c1,l1,B2B,B2C,selected_columns_df1,selected_columns_df2)

                    with open('selected.xlsx', 'rb') as f:
                        bytes = f.read()
                        b64 = base64.b64encode(bytes).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="selected.xlsx">Download data as Excel</a>'
                        st.markdown(href, unsafe_allow_html=True)
        except:
            st.error(f"An error occurred: {e}")
# Calling Main Function
if __name__ == "__main__":
    main()

