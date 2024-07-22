# Import Libraries
import pandas as pd
import numpy as np
import xml.etree.cElementTree as et
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import graphviz as graphviz
import re

# Configure Page Title and Icon
st.set_page_config(page_title='Tableau Documentor',page_icon=':smile:')

# --------- Removing Streamlit's Hamburger and Footer starts ---------
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
# --------- Removing Streamlit's Hamburger and Footer ends ---------

# Declaring custom functions

# --------- Dependent fields generator from calculation column starts ---------
#  Idea is since we know on what fields a calculation is dependent on by seeing the text inside [] in formulas,we just need to extract everything
#  which is inside square brackets and push them into new column.Also,duplicates should not be there

def dependent_fields_generator(i):
    try:
        calc_set=set()
        pattern=re.compile(r"\[(.*?)\]")
        matches=pattern.finditer(i)
        for match in matches:
            calc_set.add(str(match.group()))
        return str(calc_set).replace('[','').replace(']','').replace('{','').replace('}','').replace("'",'')
    except:
        return None
# --------- Dependent fields generator from calculation column ends ---------

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

st.header('Tableau Documentation Made Easy !!')
st.info('It takes atleast 2 minute to open a Tableau File and copy one calculation from Tableau to Excel.  \nIf you have even 10 calulations, that will take 10*2=20 minutes minimum.')
st.success('You can extract all the calculation names and formulas in seconds using this website!!')
st.markdown("---")
st.subheader('Upload your TWB file')
uploaded_file=st.file_uploader('',type=['twb'],)

if uploaded_file is not None:
    tree=et.parse(uploaded_file)
    root=tree.getroot()

    # create a dictionary of name and tableau generated name

    calcDict = {}

    for item in root.findall('.//column[@caption]'):
        if item.find(".//calculation") is None:
            continue
        else:
            calcDict[item.attrib['name']] = '[' + item.attrib['caption'] + ']'

    # list of calc's name, tableau generated name, and calculation/formula
    calcList = []

    for item in root.findall('.//column[@caption]'):
        if item.find(".//calculation") is None:
            continue
        else:
            if item.find(".//calculation[@formula]") is None:
                continue
            else:
                calc_caption = '[' + item.attrib['caption'] + ']'
                calc_name = item.attrib['name']
                calc_raw_formula = item.find(".//calculation").attrib['formula']
                calc_comment = ''
                calc_formula = ''
                for line in calc_raw_formula.split('\r\n'):
                    if line.startswith('//'):
                        calc_comment = calc_comment + line + ' '
                    else:
                        calc_formula = calc_formula + line + ' '
                for name, caption in calcDict.items():
                    calc_formula = calc_formula.replace(name, caption)

                calc_row = (calc_caption, calc_name, calc_formula, calc_comment)
                calcList.append(list(calc_row))

    # convert the list of calcs into a data frame
    data = calcList

    data = pd.DataFrame(data, columns=['Calculated Field', 'Remote Name', 'Formula', 'Comment'])

    # remove duplicate rows from data frame
    data = data.drop_duplicates(subset=None, keep='first', inplace=False)

    df=data[['Calculated Field','Formula']]
    df['Base Fields']=[str(dependent_fields_generator(i)).replace('set()','').replace(', Parameters','').replace('Parameters,','') for i in df['Formula']]

    # Showing Dataframe
    st.write(df)

    # Showing Multiple Download Options
    st.subheader("Download the above Table as : ")
    op=st.radio('Options', ["CSV File","EXCEL File"])

    # Download Functionality
    if 1==1:
        st.download_button("Download",df.to_csv(),file_name="Documentation-csv-output",mime="text/csv")
    else:
        st.download_button("Download",to_excel(df),file_name="Documentation-excel-output")


    # --------- GRAPH Functionality starts ---------
    # Data Modelling

    df1=df.copy()
    df1['Base Fields']=df1['Base Fields'].str.split(",")
    df1=df1.explode('Base Fields')
    graph = graphviz.Digraph("diag", filename="fsm.dot",format='png')
    graph.attr(rankdir="LR", size="5,5")
    #graph.attr(bgcolor="lightblue")

    # Styling
    for i in df1['Calculated Field']:
      #graph.node(i,color='red')
      graph.node(i,shape='doublecircle')

    for (i,j) in zip(df1['Calculated Field'],df1['Base Fields']):
        graph.edge(i,j)

    st.markdown('---')
    st.subheader('Variable Dependency Diagram')
    st.write("Visualize which base columns are used to make the calculated fields ")
    st.info('To see the Diagram in full width: Hover on chart, click on full view icon and print it as pdf!!')
    st.graphviz_chart(graph,use_container_width=True)
 
    # --------- GRAPH Functionality ends ---------
st.markdown('---')
st.markdown('Made with :heart: by [Sahil Choudhary](https://www.sahilchoudhary.ml/)')
