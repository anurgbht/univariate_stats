import os
import pandas as pd
import numpy as np
from datetime import datetime
import random
#######################################################################################################################
def is_datetime(col):
    global valid_formats
    global datetime_format
    no_null = col.dropna()
    no_null_list = list(no_null)
    no_null_list = [x for x in no_null_list if x != 'nan']
    no_null_list = [x for x in no_null_list if x != '0-Jan-00']
##    sample_no_null = random.sample(list(no_null),100)
    for iter_format in valid_formats:
        try:
            converted_sample = [datetime.strptime(str(x),iter_format) for x in no_null_list]
            datetime_format = iter_format
            tt = 1
            break
        except:
            tt = 0
    return tt

#######################################################################################################################
def is_numeric(col):
    no_null = col.dropna()
##    sample_no_null = random.sample(list(no_null),100)
    try:
        [float(str(x).replace(',','').replace('$','').replace('(','').replace(')','')) for x in no_null]
        tt = 1
    except:
        tt = 0
    return tt

#######################################################################################################################
def get_dtype(file):
    global categorical_cols
    global numerical_cols
    global datetime_cols
    ncol = file.shape[1]
    dtype = ['cat' for i in range(ncol)]

    # auto-detecting the column type
    for i in range(ncol):
        col = file.iloc[:,i]
        if is_numeric(col) == 1:
            dtype[i] = 'nm'
        elif is_datetime(col) == 1:
            dtype[i] = 'dt'

    # over writing the column type based on user input
    for i in range(ncol):
        if i in categorical_cols:
            dtype[i] = 'cat'
        elif i in numerical_cols:
            dtype[i] = 'nm'
        elif i in datetime_cols:
            dtype[i] = 'dt'
    return dtype

#######################################################################################################################
def get_num_info(col,i):
    if is_numeric(col):
        no_null = pd.Series([float(str(x).replace(',','').replace('$','').replace('(','').replace(')','')) for x in col.dropna()])
    ##    print(no_null)
        try:
            temp = ['numerical',col.name,i,len(col),col.isnull().sum(),(100*(col.isnull().sum())/(len(col))),len(no_null.unique()),no_null.mean()
                    ,no_null.min(),no_null.quantile(q=0.01),no_null.quantile(q=0.1),no_null.quantile(q=0.05),no_null.quantile(q=0.25),
                    no_null.quantile(q=0.5),no_null.quantile(q=0.75),no_null.quantile(q=0.9),no_null.quantile(q=0.95)
                    ,no_null.quantile(q=0.99),no_null.max()]
        except:
            if (len(no_null.unique()) == 1):
                unique_el = float(no_null.unique()[0])
                temp = ['numerical',col.name,i,len(col),col.isnull().sum(),(100*(col.isnull().sum())/(len(col))),unique_el,unique_el, \
                               unique_el,unique_el,unique_el,unique_el,unique_el,unique_el]
            else:
                temp = ['numerical',col.name,i,'error','error','error','error','error','error','error','error','error','error','error']
        print(temp)
    else:
        temp = ['numerical',col.name,i,'error','error','error','error','error','error','error','error','error','error','error']
        raise ValueError('Column coercion did not work for this numerical column. Please check the coercion and try again. The column name and number are printed above.')
    return temp

#######################################################################################################################
def get_cat_info(col,i):
    # categorical coercion should always work !!
    global n_freq
    no_null = col.dropna()
    try:
        temp = ['categorical',col.name,i,len(col),col.isnull().sum(),(100*(col.isnull().sum())/(len(col))),len(no_null.unique()),len(no_null) - no_null.value_counts().iloc[:n_freq].sum()]
        temp2 = [[x,y] for (x,y) in zip(no_null.value_counts().index[:n_freq].tolist(),no_null.value_counts().iloc[:n_freq].tolist())]
        for i in temp2:
            temp.extend(i)
    except:
        temp = ['categorical',col.name,i,'error','error','error','error','error','error','error','error','error','error','error']
    print(temp)
    return temp

#######################################################################################################################
def get_date_info(col,i):
    global datetime_format
    if is_datetime(col):
        no_null = col.dropna()
        no_null_list = list(no_null)
        no_null_list = [x for x in no_null_list if x != 'nan']
        no_null_list = [x for x in no_null_list if x != '0-Jan-00']
        no_null_list = [datetime.strptime(str(x),datetime_format) for x in no_null_list]
        try:
            temp = ['date',col.name,i,len(col),len(col)-len(no_null_list),(100*(len(col)-len(no_null_list))/(len(col))),min(no_null_list),max(no_null_list)]
        except:
            temp = ['date',col.name,i,'error','error','error','error','error']
    else:
        temp = ['date',col.name,i,'error','error','error','error','error']
        raise ValueError('Column coercion did not work for this datetime column. Please check the coercion and try again. The column name and number are printed above.')
    print(temp)
    return temp

#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

os.chdir("D:/OneDrive - Tata Insights and Quants, A division of Tata Industries/Confidential/Projects/Steel/LD2 BOF/data/jan17-oct17/static/")

# these are user inputs
global n_freq
global valid_formats
global categorical_cols
global numerical_cols
global datetime_cols
global datetime_format
n_freq = 5
valid_formats = ["%Y-%m-%d","%d.%m.%y","%d.%M.%Y","%d-%m-%Y %I:%M","%d-%m-%y %I:%M","%d-%m-%Y %I:%M:%S %p",
                 "%Y-%m-%d %I:%M:%S %p","%d-%m-%y","%Y-%m-%d %I:%M:%S","%Y-%m-%d %H:%M:%S","%Y%m","%m/%d/%Y %H:%M",
                 "%d-%m-%Y","%m/%d/%Y","%d-%m-%Y %I:%M:%S %p","%d-%m-%Y %H:%M","%d-%m-%y %H:%M","%d-%b-%y","%b-%y"
                 ]

categorical_cols = []
numerical_cols = []
datetime_cols = []

##file_list=['205 BGG_ FY 2016.XLSX']
file_list = os.listdir()
writer = pd.ExcelWriter('BOF_static.xlsx')
for file_name in file_list:
    print(file_name)
    # checking the file type using the extension and reading accordingly
    if ((file_name.split('.')[-1].lower() == 'xlsx') or (file_name.split('.')[-1].lower() == 'xls')):
        file = pd.read_excel(file_name)
    elif (file_name.split('.')[-1].lower() == 'csv'):
        file = pd.read_csv(file_name)
    print(file.shape)
    dtype = get_dtype(file)
    temp_cat = []
    temp_num = []
    temp_date = []
    for i in range(len(dtype)):
    ##    print(dtype[i])
        tt = file.iloc[:,i]
        print('Executing for column number : ',i,' Name of the column : '+tt.name)
        if dtype[i] == 'cat':
            temp_cat.append(get_cat_info(tt,i))
        elif dtype[i] == 'nm':
            temp_num.append(get_num_info(tt,i))
        elif dtype[i] == 'dt':
            temp_date.append(get_date_info(tt,i))

    if len(temp_num)>0:
        temp_num = pd.DataFrame(temp_num)
        temp_num.columns = ['type','col_name','col_number','count','missing','missing perc','number of unique','mean',
                            'minimum','1st percentile','5th percentile','10th percentile','25th percentile','50th percentile',
                            '75th percentile','90th percentile','95th percentile','99th percentile','maximum']
##        temp_num.to_csv(file_name.split('.')[0]+'_num.csv',index=False)
        temp_num.to_excel(writer,sheet_name = file_name.split('.')[0]+'_num')

    if len(temp_date)>0:
        temp_date = pd.DataFrame(temp_date)
        temp_date.columns = ['type','col_name','col_number','count','missing','missing perc','range from','range to']
##        temp_date.to_csv(file_name.split('.')[0]+'_date.csv',index=False)
        temp_date.to_excel(writer,sheet_name = file_name.split('.')[0]+'_date')

    if len(temp_cat)>0:
        temp_cat = pd.DataFrame(temp_cat)
        tt2 = ['type','col_name','col_number','count','missing','missing perc','number of unique values','count of others']
        col_name_temp=[]
        col_temp = ['value','freq']
        n_tt = int((temp_cat.shape[1]-7)/2)
        for i in range(n_tt):
            col_name_temp.extend(col_temp)
        tt2.extend(col_name_temp)
        temp_cat.columns = tt2
##        temp_cat.to_csv(file_name.split('.')[0]+'_cat.csv',index=False)
        temp_cat.to_excel(writer,sheet_name = file_name.split('.')[0]+'_cat')

writer.save()
