#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import packages
import pandas as pd
import sys


# ## P/L Statement
# ---

# In[2]:


# initialize the inputs (py, cy, and currency unit)
def init_inputs():
    print('Initializing the inputs...')
    py = input('Enter the prior / last year (e.g., FY20): ')
    cy = input('Enter the current year (e.g., FY21): ')
    currency = input('Enter the currency unit (e.g., $, CHF, GBP). If using the 3-letter currency code, such as GBP, enter a space after the 3-letters.: ')
    file_path = input('Paste your file path. Be sure to include the file name at the end (e.g., C:\\Users\\Jed\\Documents\\mydata.xlsx): ')
    return(py, cy, currency, file_path.replace('\\', "\\\\"))


# In[3]:


# load test data
def load_data(file_path):
    print('Loading the data...')
    df_l2 = pd.read_excel(file_path)
    df_l1 = pd.pivot_table(df_l2, index = ['L1 Sort', 'L1'], values = ['PY', 'CY'], aggfunc = sum)
    return(df_l2, df_l1)


# In[4]:


# add YoY delta
def cy_change(df):
    print('Calculating the changes...')
    df['delta'] = df['CY'] - df['PY']
    return(df)


# In[5]:


# categorize as increase or decrease
def inc_dec(df):
    print('Categorizing increases, decreases, or consistencies...')
    for row in range(len(df.index)):
        if abs(df['CY'].iloc[row]) > abs(df['PY'].iloc[row]):
            df['inc_dec'] = 'increased'
        elif abs(df['CY'].iloc[row]) == abs(df['PY'].iloc[row]):
            df['inc_dec'] = 'remained consistent'
        else:
            df['inc_dec'] = 'decreased'
    return(df)


# In[6]:


# initialize outputs
def init_outputs():
    print('Initializing outputs...')
    output = []
    all_drivers = []
    sentences = []
    result = ''
    return(output, all_drivers, sentences, result)


# In[7]:


# base of the write-up
def write_up_base_pl(df_l1, df_l2, output, currency):
    print('Drafting the base of the write-up...')
    for i in range(len(df_l1)):
        output.append(['[{}] {} by {}{} from {}{} in PY to {}{} in CY driven by '.format(
            df_l1.iloc[i]['L1'], 
            df_l1.iloc[i]['inc_dec'], 
            currency,
            '{:,}'.format(round(df_l1.iloc[i]['delta'])),  
            currency,
            '{:,}'.format(round(df_l1.iloc[i]['PY'])),  
            currency,
            '{:,}'.format(round(df_l1.iloc[i]['CY'])))])
    return(output)


# In[8]:


# drivers write ups
def write_up_drivers_pl(df_l1, df_l2, all_drivers, currency):
    print('Drafting the drivers in the write-up...')
    for level_1 in df_l2['L1'].unique():
        sub_df = df_l2[df_l2['L1'] == level_1]
        if len(df_l2[df_l2['L1'] == level_1]) > 1:
            drivers = []
            for i in range(len(sub_df)):
                drivers.append('[{}] ({}{})'.format(
                    sub_df.iloc[i]['L2'], 
                    currency, 
                    '{:,}'.format(round(sub_df.iloc[i]['delta']))))
            all_drivers.append(drivers)
        elif len(sub_df) == 1:
            all_drivers.append('[OPEN]')
    return(all_drivers)


# In[9]:


# intermediate length check
def intermediate_check(all_drivers, df_l1):
    print('Process check / reconciliation...')
    continue_response = ''
    if len(all_drivers) == len(df_l1):
        print('No errors found, continuing...')
    else:
        print('WARNING: Check resulted in a mismatch in legnths between the write-up and the dataset. Please inspect the result to ensure proper coverage. To continue enter \'y\' ')
        continue_response = input('Continue? (y/n): ')
        if continue_response == 'y':
            print('User entered y, continuing...')
        else:
            print('User entered n, exiting...')
            sys.exit()
    return


# In[10]:


# join drivers
def join_drivers(df_l1, all_drivers, output):
    print('Joining the drivers...')
    for i in range(len(df_l1)):
        if all_drivers[i] == '[OPEN]':
            output[i].append('[OPEN]')
        elif len(all_drivers[i]) == 2:
            output[i].append(' and '.join(all_drivers[i]))
        elif len(all_drivers[i]) > 2:
            output[i].append(', '.join(all_drivers[i]))
    return


# In[11]:


# join sentences
def join_sentences(sentences, output):
    print('Joining the sentences...')
    for i in range(len(output)):
        sentences.append(''.join(output[i]))
    return


# In[12]:


# join result
def join_result(sentences):
    print('Joining the result...')
    result = '. \n'.join(sentences) + '.'
    return(result)


# In[13]:


# replace CY and PY with actual periods
def replace_periods(result, py, cy):
    print('Labeling the periods...')
    result_replaced = result.replace('PY', py).replace('CY', cy)
    return(result_replaced)


# ## Balance Sheet
# ---

# In[14]:


# base of the write-up
def write_up_base_bs(df_l1, df_l2, output, currency):
    print('Drafting the base of the write-up...')
    for i in range(len(df_l1)):
        output.append(['[{}] of {}{} in CY relates to '.format(
            df_l1.iloc[i]['L1'],    
            currency,
            '{:,}'.format(abs(df_l1.iloc[i]['CY'])))])
    return(output)


# In[15]:


# drivers write ups
def write_up_drivers_bs(df_l1, df_l2, all_drivers, currency):
    print('Drafting the drivers in the write-up...')
    for level_1 in df_l2['L1'].unique():
        sub_df = df_l2[df_l2['L1'] == level_1]
        if len(df_l2[df_l2['L1'] == level_1]) > 1:
            drivers = []
            for i in range(len(sub_df)):
                drivers.append('[{}] ({}{})'.format(
                    sub_df.iloc[i]['L2'], 
                    currency, 
                    '{:,}'.format(round(sub_df.iloc[i]['CY']))))
            all_drivers.append(drivers)
        elif len(sub_df) == 1:
            all_drivers.append('[OPEN]')
    return(all_drivers)


# ## Run the Script
# ---

# In[16]:


# define main function
def main():
    # get financial statement type
    financial_statement = input('Is this an income statement or balance sheet? Enter 1 for income statement or 2 for balance sheet: ')
    
    # income statement
    if financial_statement == '1':
        print('Income Statement')
        py, cy, currency, file_path = init_inputs()
        df_l2, df_l1 = load_data(file_path)
        cy_change(df_l1)
        cy_change(df_l2)
        inc_dec(df_l1)
        df_l1 = df_l1.reset_index()
        output, all_drivers, sentences, result = init_outputs()
        output = write_up_base_pl(df_l1, df_l2, output, currency)
        all_drivers = write_up_drivers_pl(df_l1, df_l2, all_drivers, currency)
        intermediate_check(all_drivers, df_l1)
        join_drivers(df_l1, all_drivers, output)
        join_sentences(sentences, output)
        result = join_result(sentences)
        result = replace_periods(result, py, cy)
        print('----------')
        print(result)
    # balance sheet
    elif financial_statement == '2':
        print('Balance Sheet')
        py, cy, currency, file_path = init_inputs()
        df_l2, df_l1 = load_data(file_path)
        df_l1 = df_l1.reset_index()
        output, all_drivers, sentences, result = init_outputs()
        output = write_up_base_bs(df_l1, df_l2, output, currency)
        all_drivers = write_up_drivers_bs(df_l1, df_l2, all_drivers, currency)
        intermediate_check(all_drivers, df_l1)
        join_drivers(df_l1, all_drivers, output)
        join_sentences(sentences, output)
        result = join_result(sentences)
        result = replace_periods(result, py, cy)
        print('----------')
        print(result)
    # invalid entry
    else:
        print('Invalid entry. Exiting...')
        sys.exit()


# In[18]:


# C:\Users\Jed\iCloudDrive\Documents\Learn\Python\Plabs\Plabs_PL.xlsx
# C:\Users\Jed\iCloudDrive\Documents\Learn\Python\Plabs\test_data_bs.xlsx

main()

