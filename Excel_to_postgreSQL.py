import pandas as pd
from sqlalchemy import create_engine
import os
import psycopg2

# creating connection
engine = create_engine("postgresql+psycopg2://postgres:10aabg58@localhost:5432/Electrode_experiments2")

# starting by this code to test
# with pd.ExcelFile(r'C:\Users\mou95504\Desktop\Test\InMiTro_Kathode_1.xlsx') as xls:
#    df = pd.read_excel(xls, sheet_name='2.experiments', index_col=None)
#    df.to_sql(name = 'experiments', con= engine, if_exists='append', index= False)

def sheets(data, file):
    if(data=='1.projects'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='projects', con=engine, if_exists= 'append', index= False) # name is table name 

    elif(data=='2.experiments'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='experiments', con=engine, if_exists= 'append', index= False) # name is table name
    
    elif(data=='3.meas.steps'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='measurement_steps', con=engine, if_exists= 'append', index= False) # name is table name 
    
    elif(data=='4.Proces.steps'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='processing_steps', con=engine, if_exists= 'append', index= False) # name is table name 

    elif(data=='5.mater.add.steps'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='material_addition_steps', con=engine, if_exists= 'append', index= False) # name is table name  

    elif(data=='6.slurry.mater.'):
        # df=pd.read_excel(file)
        df = pd.read_excel(file, sheet_name=data, index_col=None)
        df.to_sql(name='slurry_materials', con=engine, if_exists= 'append', index= False) # name is table name      
        

# with pd.ExcelFile(r'C:\Users\mou95504\Desktop\Test\InMiTro_Kathode_1.xlsx') as xls:
#     for sheet_name in xls.sheet_names:
#         sheets(sheet_name, xls)

# Folder Path
path = r'C:\Users\mou95504\Desktop\Test'
# Change the directory
os.chdir(path)

for file in os.listdir():
    # print("hi")
    # Check whether file is in text format or not
    if file.endswith(".xlsx"):
        file_path = f"{path}\{file}"
        with pd.ExcelFile(file_path) as xls:
            for sheet_name in xls.sheet_names:
                sheets(sheet_name, xls)

