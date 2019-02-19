import numpy as np
import pandas as pd
import dask.dataframe as dd
import scipy.stats as stats
import matplotlib.pyplot as plt
import sklearn
import statsmodels
import statsmodels.api as sm
import statsmodels.formula.api as smf   #FOR USING 'R'-STYLE FORMULAS FOR REGRESSIONS
import imaplib, email, os
import configparser
import datetime
from scipy.stats import norm
from sklearn.preprocessing import StandardScaler
from scipy import stats
from pandas import DataFrame
from pandas import Series
import warnings
warnings.filterwarnings('ignore')
import seaborn as sns
sns.set_style("whitegrid")
sns.set_context("poster")
color = sns.color_palette()

# special matplotlib argument for improved plots
from matplotlib import rcParams

pd.set_option("display.max_rows",None)
pd.set_option("display.max_columns", None)
pd.set_option('display.float_format', lambda x: '%.2f' % x)

import imaplib, email, os
import configparser
import datetime
import shutil
import os
import glob

import openpyxl
import pprint


path =r'\\prod-app02\ConvenienceCheckApproval\WellsFargoFiles' # use your path
allFiles = glob.glob(path + "/*.xlsx")
bank = pd.DataFrame()
list_ = []
for file_ in allFiles:
    bank = pd.read_excel(file_,index_col=None, header=0,
                         encoding="ISO-8859-1", error_bad_lines=False)
    list_.append(bank)
bank = pd.concat(list_)
bank.columns = bank.columns.str.strip().str.lower().str.replace(' ', '_')
bank.columns = bank.columns.str.strip().str.lower().str.replace('-', '_')

bank['customer_ref_no'] = bank['customer_ref_no'].astype(str) 
bank['customer_ref_no'] = bank['customer_ref_no'].apply(lambda x: '{0:0>10}'.format(x))

updates = dict(zip(bank.customer_ref_no, bank.as_of_date))

for i in os.listdir(os.chdir(r'M:\2019 Programs\Master Spreadsheets')):
    if i.endswith(".xlsx"):
        wb = openpyxl.load_workbook(i)
        sheet = wb.get_sheet_by_name('Setupheader')
        for rowNum in range(2, sheet.max_row):  # skip the first row
            checkNumber = sheet.cell(row=rowNum, column=19).value
            if checkNumber in updates:
                sheet.cell(row=rowNum, column=20).value = updates[checkNumber]
        wb.save(i)

for i in os.listdir(os.chdir(r'M:\2019 Programs')):
    if i.endswith("TEST.xlsx"):
        wb = openpyxl.load_workbook(i)
        sheet = wb.get_sheet_by_name('Setupheader')
        for rowNum in range(2, sheet.max_row):  # skip the first row
            checkNumber = sheet.cell(row=rowNum, column=19).value
            if checkNumber in updates:
                sheet.cell(row=rowNum, column=20).value = updates[checkNumber]
        wb.save(i)