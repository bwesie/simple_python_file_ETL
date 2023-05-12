# import modules
import pandas as pd
import glob
import os
import shutil

shutil.rmtree('insertFilePathHere')

# Providing the folder path
origin = r'insertFilePathHere'
target = r'insertFilePathHere'

# Fetching the list of all the files
files = os.listdir(origin)

shutil.copytree(origin, target)

os.chdir(r'insertFilePathHere')

# specifying the path to excel files
path = r'insertFilePathHere'

# excel files in the path
file_list = glob.glob(path + r'\*.xlsx')

# list of excel files we want to merge.
# pd.read_excel(file_path) reads the excel data into pandas dataframe.
excl_list = []

for file in file_list:
    excl_list.append(pd.read_excel(file))

# create a new dataframe to store the merged excel file.
excl_merged = pd.DataFrame()

for excl_file in excl_list:
    # appends the data into the excl_merged dataframe.
    excl_merged = excl_merged.append(
        excl_file, ignore_index=True)

# exports the dataframe into excel file with specified name.
excl_merged.to_excel('insertFinalReportNameHere.xlsx', index=False)
