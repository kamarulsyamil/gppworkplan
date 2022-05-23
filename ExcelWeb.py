import pandas as pd
import streamlit as st
import json

st.set_page_config(page_title='Factories Workplans Consolidated')
st.header('GPP Workplan (Shift times)')
st.subheader('Updated on :')

# LOAD EXCEL
# read config file
with open(r"tool_config.json") as config_file:
    config = json.load(config_file)
    file_dir = config['file_dir']

excel_file = file_dir['main_excel']

sheet_name = 'Workplans'

df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='B:K', header=7)
df2 = df.fillna('')
st.write(df2.astype(str))

option = st.selectbox(
    'Choose factory',
    ('CCC4', 'CCC2', 'CCC6', 'APCC', 'ICC', 'EMFP', 'BRH1'))

st.write('You selected:', option)

if option == 'CCC4':
    st.write(df2.iloc[0:7, :].astype(str))

elif option == 'CCC2':
    st.write(df2.iloc[9:15, :].astype(str))

elif option == 'CCC6':
    st.write(df2.iloc[17:18, :].astype(str))

elif option == 'APCC':
    st.write(df2.iloc[25:31, :].astype(str))

elif option == 'ICC':
    st.write(df2.iloc[33:39, :].astype(str))

elif option == 'EMFP':
    st.write(df2.iloc[41:42, :].astype(str))

elif option == 'BRH1':
    st.write(df2.iloc[49:56, :].astype(str))
