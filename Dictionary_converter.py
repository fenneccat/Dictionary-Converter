#!/usr/bin/env python
# coding: utf-8

# In[1]:


bookname = input('Put book/movie name: ')


# In[2]:


import os
import os.path
from os import path

from datetime import datetime

if not os.path.exists(bookname):
    print('[{}] Create folders and files that you need'.format(datetime.datetime.now().time()))
    os.makedirs(bookname)
    pathprefix = './'+bookname+'/'
    files = [pathprefix+bookname+'_new.xlsx', pathprefix+bookname+'_DB.csv']
    for file in files:
        f = open(file, 'w')
        f.close()


# In[3]:


import pandas as pd


try:
    newdata = pd.read_excel('./'+bookname+'/'+bookname+'_new.xlsx',header = None, names = ['Eng','Kor'])
    print('[{}] Reading {}_new.xlsx file complete'.format(datetime.now().time(), bookname))
except:
    print('{}_new.xlsx file doesn\'t exist'.format(bookname))
    print('[{}] Creating {}_new.xlsx file...'.format(datetime.now().time(), bookname))
    f = open('./'+bookname+'/'+bookname+'_new.xlsx', 'w')
    newdata = pd.DataFrame()
    f.close()

try:
    DB=pd.read_csv('./'+bookname+'/'+bookname+'_DB.csv')
    print('[{}] Loading DB file complete'.format(datetime.now().time(), bookname))
except:
    print('{}_DB.csv file doesn\'t exist'.format(bookname))
    print('[{}] Creating {}_DB.csv file...'.format(datetime.now().time(), bookname))
    f = open('./'+bookname+'/'+bookname+'_DB.csv', 'w')
    DB = pd.DataFrame()
    f.close()


# In[4]:


print('[{}] Building Dictionary...'.format(datetime.now().time()))

from collections import Counter
engdict = dict()
engcounter = Counter()
for i, row in DB.iterrows():
    engdict[row.Eng] = row.Kor
    engcounter[row.Eng] = row.cnt


# In[5]:


seen = set()
newwords = set(newdata.Eng)
for i, row in newdata.iterrows():
    if row.Eng in engdict:
        engcounter[row.Eng] += 1
        seen.add(row.Eng)
    else:
        engdict[row.Eng] = row.Kor
        engcounter[row.Eng] = 1


# In[6]:


englst = []
korlst = []
cntlst = []
for eng, kor in engdict.items():
    englst.append(eng)
    korlst.append(kor)
    cntlst.append(engcounter[eng])


# In[7]:


data = pd.DataFrame({'Eng': englst,'Kor': korlst,'cnt': cntlst})


# In[8]:


print('[{}] Saving new DB...'.format(datetime.now().time()))
data.to_csv('./'+bookname+'/'+bookname+'_DB.csv',index=False,encoding='utf-8-sig')


# In[9]:


nondf = pd.DataFrame()


# In[10]:


print('[{}] Clearing {}_new.xlsx file...'.format(datetime.now().time(), bookname))
nondf.to_excel('./'+bookname+'/'+bookname+'_new.xlsx', index = False)


# In[11]:


data['startletter'] = data.Eng.str[0].str.lower()


# In[13]:


from docx import Document
from docx.shared import Inches
from docx.shared import Cm
from docx.shared import RGBColor
import tqdm
 
print('[{}] Writing on Vocabulary MS Words file...'.format(datetime.now().time()))
    
document = Document()

height = 0.3

print('[{}] most frequent words part...'.format(datetime.now().time()))
 
document.add_heading(bookname+' Vocabulary', 0)
data = data.sort_values('cnt', ascending=False)
p = document.add_heading('Frequently confused words',level=1)

table_fq = document.add_table(data.shape[0]+1, data.shape[1]-1)

hdr_cells = table_fq.rows[0].cells

# add the header rows.
for j in range(data.shape[-1]-1):
    run = hdr_cells[j].paragraphs[0].add_run(data.columns[j])
    run.bold = True
    run.underline = True
    #table_fq.cell(0,j).text = data.columns[j]

# add the rest of the data frame
check = False
newreg = False
for i in tqdm.tqdm(range(data.shape[0])):
    check = False
    newreg = False
    if str(data.values[i,0]) in seen:
        check = True
    if str(data.values[i,0] in newwords):
        newreg = True
    for j in range(data.shape[-1]-1):
        cell = table_fq.cell(i+1,j)
        cell.text = str(data.values[i,j])
        if check:
            run = cell.paragraphs[0].runs[0]
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        elif newreg:
            run = cell.paragraphs[0].runs[0]
            run.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
        

for row in table_fq.rows:
    row.height = Cm(height)


data["lower"] = data["Eng"].str.lower()
data.sort_values(['lower'], axis=0, ascending=True, inplace=True)
del data['lower']
del data['cnt']


# In[14]:


print('[{}] word dictionary part...'.format(datetime.now().time()))
p = document.add_heading('Word Dictionary',level=1)
data_group_letter = data.groupby('startletter')
for i in range(26):
    ch = chr(ord('a')+i)
    if ch in data_group_letter.groups.keys():
        p = document.add_heading(ch.upper(),level=2)
        groupdata = data_group_letter.get_group(ch)

        table_dict = document.add_table(groupdata.shape[0], groupdata.shape[1]-1)

        # add the rest of the data frame
        for i in range(groupdata.shape[0]):
            for j in range(groupdata.shape[-1]-1):
                table_dict.cell(i,j).text = str(groupdata.values[i,j])
        
        for row in table_dict.rows:
            row.height = Cm(height)

document.save('./'+bookname+'/'+bookname+'_Dictionary.docx') # Save document


# In[ ]:




