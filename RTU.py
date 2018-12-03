
# coding: utf-8

# In[1]:


import csv
import xlrd

workbook = xlrd.open_workbook('C:\Users\talwe\OneDrive\Desktop\Linworth.xlsx')
for sheet in workbook.sheets():
    with open('{}.csv'.format(sheet.name), 'w') as f:
        writer = csv.writer(f)
        writer.writerows(sheet.row_values(row) for row in range(sheet.nrows))

print ("CSV converted")


# need to modify this to ignore 1-9 rows and convert .xlsm to csv from 10th row
# since headers are merged in row 1 and row 2, the csv has two lines which will give wrong starting index name refrence.


# In[2]:


import pandas as pd


# In[3]:


xl_file = pd.ExcelFile('C:\\Users\\talwe\\OneDrive\\Desktop\\Linworth.xlsx')
dfs = {sheet_name: xl_file.parse(sheet_name) 
          for sheet_name in xl_file.sheet_names}


# In[4]:


dfs['Digital Inputs (VVO)']


# In[5]:



dfs.keys()


# In[ ]:


dfs['Digital Inputs (VVO)']['Unnamed: 2'][9:]


# In[ ]:


feeders = dfs['Inside Fence Addressing']['Unnamed: 4'][4:-5].tolist()
feeders


# In[ ]:


f_nums = list(set([entry[-5:-1] for entry in feeders]))
f_nums


# In[ ]:


dfs['Analog Inputs ']


# In[ ]:


analogs = dfs['Analog Inputs ']['Unnamed: 2'][9:-8]
#analogs = dfs['Analog Inputs ']['Unnamed: 2'][9:-8].tolist()


# In[ ]:


#analogs[0]


# In[ ]:


analogs.str.split()


# In[ ]:


"""
From: 8-09-13.2kV CB8 F-4808 Ph-a MW
To: F4808.MTR.0.P.A
"""


# In[ ]:


# Split
# -3: If consists feeder number
# -2: last letter, after ph-
# -1: MW


# In[11]:


mws = [entry for entry in analogs if (entry.endswith("MW") and "F" in entry and "XF" not in entry)]


# In[ ]:


mws


# In[ ]:


amps = [entry for entry in analogs if (entry.endswith("Amps") and "F" in entry)]
#amps = [entry for entry in analogs if (entry.endswith("Amps") and "F" in entry and "REG" not in entry)]


# In[ ]:


amps


# In[ ]:


#[words for segments in amps for words in segments.split()]


# In[ ]:


my_list =["REG", "CAP", "LVM", "CB", "MLR", "LTC"]
my_list


# In[ ]:



# Match names
for element in amps:
     m = re.match("my_list", element)
     if m:
        print(m)


# In[69]:


analogs = dfs['Analog Inputs ']['Unnamed: 2'][54:173]
analogs
with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    print(analogs)


# In[30]:


d = {'Ph-a':'A', 'Ph-b':'B', 'Ph-c':'C', 'Amps':'I', 'V':'V', 'MW':'P', 'MVAR':'Q', 'KVA':'S', 'Amps':'I', 'Neutral Amps':'I.N' }





# In[ ]:


# Acronyms: CI=caseinsensitive

# Remove Till First Space
# PHASE(CI) dictionary: -a=A, -b=B, -c=C, 3-=3, -an=AN, -bn=BN, -cn=CN, Neutral=N
# UNIT(CI): Amps=I, kV=V, (MW/KW)=P, MVar(CI)=Q
# DEVICETYPE(CI): 
#   CB(DIGITS)=MTR
#   BUS (PHASE=A|B|C)=MTR
# FEEDERID: F-NUMBER|FNUMBER=FNUMBER, unless BUS (raise boolean)

# For MTR: FEEDERID.DEVICETYPE.0.UNIT.PHASE


# In[70]:


#ToDo: Other case variants
phases_ = {'Ph-a':'A', 'Ph-b':'B', 'Ph-c':'C', '3-ph':'3PH', 'Ph-an':'AN', 'Ph-bn':'BN', 'Ph-cn':'CN', 'Neutral':'N'}
units_ = {'Amps':'I', 'kV':'V', 'MW':'P', 'KW':'P', 'MVar':'Q', 'MVAR':'Q'}
devicetype_start_ = {'CB':'MTR'}






# In[71]:


import re


# In[72]:


rtu_meters = []
for name in analogs:
    parts = name.split()
    unit = units_[parts[-1]]
    phase = phases_[parts[-2]]
    if parts[1] == 'Bus':
    
        feeder = rtu_meters[0].split('.')[0]
        devicetype = rtu_meters[0].split('.')[1]
    else:
        devicetype = devicetype_start_[parts[1][:2]]
        feeder = 'FR' + re.findall('\d+', parts[-3])[0]
    rtu_meters.append(feeder + '.' + devicetype + '.' + '0' + '.' + unit + '.'+ phase)


# In[73]:


rtu_meters

