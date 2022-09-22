import os
import pandas as pd
import xlwings as xw
from IPython.display import display


# conf
path_of_dir = "C:\C-STEP FOLDER"


#final all the data frame
whole_df =[]

#function collect and store all the dataframe in the xlsx or xls sheet.
def _process(filename) -> list:
    """
    this function does ....
    parms...
    outputs...
    """
    
    wb = xw.Book(filename)
    df_list = []

    for i in wb.sheets:
        df = i.used_range.options(pd.DataFrame, index=False, header=True).value
        df_list.append(df)  

    whole_df.append(df_list)
    wb.close()

    return df_list
        

# get all the file of xlsx in the current directory
x_files = [os.path.join(root, name)
             for root, dirs, files in os.walk(path_of_dir)
             for name in files
             if name.endswith((".xls",".xlsx"))]


#iterator of files to process xlsx or xls call _process function
for f in x_files:
    display(_process(f))