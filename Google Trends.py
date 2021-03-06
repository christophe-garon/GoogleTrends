#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pytrends.request import TrendReq
from scipy.stats import linregress
import time
import openpyxl
from openpyxl import load_workbook
from datetime import date
import math
import pandas as pd
from string import ascii_uppercase
import numpy as np

pytrends = TrendReq(hl='en-US', tz=360, timeout=(10, 25))


# In[2]:


#Import the excel doc of keywords
terms_df = pd.read_excel('GoogleTrendTerms.xlsx', index_col=None)


# In[3]:


#Get the terms and monthly search volumes if they exists
terms = list(terms_df['Terms'])

try:
    volume = list(terms_df['Volume'])
except:
    volume = [0 for num in range(0,len(terms))]


# In[4]:


#Loops through the list of keywords and creates pandas df out of the scrape_google function output

def get_trends(terms):
    for i in range(0,len(terms)):
        if i == 0:
            trends = scrape_google(terms[i:i+1])
        else:
            trends = pd.concat([trends, scrape_google(terms[i:i+1])], axis=1)
    return trends


# In[5]:


#All categories: https://github.com/pat310/google-trends-api/wiki/Google-Trends-Categories

def scrape_google(terms):
    pytrends.build_payload(terms, cat=0, timeframe='2018-01-01 2020-01-01', geo='US', gprop='')
    trends = pytrends.interest_over_time()
    try:
        trends = trends.drop(columns=['isPartial'])
    except:
        pass
    
    #time.sleep(1)
    return trends


# In[6]:


#Scrape the list of keywords and 
trends_df = get_trends(terms)
trends = trends_df.copy()
trends


# In[7]:


#Create a week column and add to DataFrame
week_number = []
i = 1
for row in range(0,len(trends)):
    week_number.append(i)
    i+=1
    
trends.insert(0, 'week number', week_number)


# In[8]:


#make the indexes strings
trends.index = trends.index.astype(str) 


# In[9]:


#Should prob check that the stardard error is exceptable


# In[10]:


#Calc the percent change from year one to year 2
def get_year_avg(col):
    year_avgs = [sum(list(trends[col][0:52]))/12,sum(list(trends[col][52:105]))/12]
    if year_avgs[0] == 0:
        change = 100
    else:
        change = ((year_avgs[1] - year_avgs[0])/year_avgs[0])*100
    return change


# In[11]:


#Create list of search volume metrics to create trend_stats df

week = list(trends['week number'])
percent_change = []
yearly_change = []
standard_error = []

for col in list(trends.columns)[1:]:
    try:
        percent_change.append(linregress(week, list(trends[col]))[0] * 100)
        standard_error.append(linregress(week, list(trends[col]))[4])
        yearly_change.append(get_year_avg(col))
    except:
        percent_change.append(0)
        standard_error.append(0)
        yearly_change.append(0)


# In[12]:


#Construct trend_stats df
data = {
    "Search Term":list(trends.columns)[1:],
    "Average Change %":percent_change,
    "Yearly Change %": yearly_change,
    "Standard Error":standard_error
}

trend_stats = pd.DataFrame(data) 


# In[13]:


trends = trends.drop(columns=['week number'])

#Transpose the trends and make a copy. Makes calculations more intuitive
sv_trends = trends.copy().transpose()
trends = trends.transpose()


# In[14]:


volumes_col = []
for i in range(0, len(terms)):
    if terms[i] in list(trends.index):
        volumes_col.append(volume[i])

try:
    trends.insert(0, "avg monthly volume", volumes_col)
except:
    pass

trends


# In[15]:


for i in range(0,len(sv_trends.index)):
    sv_trends[i:i+1] = sv_trends[i:i+1].mul((volume[i]/4)/50)
    
sv_trends.insert(loc=0, column='Trend Lines', value=['' for i in range(sv_trends.shape[0])])


# In[16]:


def get_col(col_number):
    cols = ascii_uppercase
    if col_number <= 26:
        return cols[col_number-1]
    else:
        col = col_number//26
        col_rem = col_number%26
        if col_number/26 != col:
            col2 = col_number%26
            col =  cols[col-1] + cols[col2-1]
        else:
            col = cols[col-2] + cols[col_rem-1]
        
        return col


# In[17]:


today = date.today()

writer = pd.ExcelWriter("GoogleTrends({}).xlsx".format(today), engine='xlsxwriter')
trends.to_excel(writer, sheet_name='Trends', index=True)
sv_trends.to_excel(writer, sheet_name='Volume Trends', index=True)
trend_stats.to_excel(writer, sheet_name='Trend Stats', index=False)


wb = writer.book
stats_sheet = writer.sheets['Trend Stats']
trends_sheet = writer.sheets['Trends']
vol_sheet = writer.sheets['Volume Trends']

#trends sheet formatting
trends_sheet.set_column(2, len(trends.columns), 10)
trends_sheet.set_column(0, 1, 25)
trends_sheet.set_column(1, 2, 18)

for row in range(2,len(trends)+2):
    trends_sheet.conditional_format('C{}:{}{}'.format(row, get_col(len(trends.columns)+1), row), {'type': '3_color_scale'})


#volume sheet formatting    
vol_sheet.set_column(2, len(sv_trends.columns), 10)
vol_sheet.set_column(0, 1, 25)


for row in range(0,len(sv_trends)):
    vol_sheet.add_sparkline('B{}'.format(row+2), {'range': 'Trends!{}{}:{}{}'.format(get_col(3),row+2, get_col(len(trends.columns)), row+2)})
    vol_sheet.set_row(row+2, 25)


for row in range(0,len(sv_trends)+2):
    vol_sheet.conditional_format('B{}:{}{}'.format(row, get_col(len(trends.columns)+1), row), {'type': '3_color_scale'})

    
#stat sheet formatting
stats_sheet.set_column('A:A', 30)
stats_sheet.set_column('E:E', 100)
stats_sheet.set_column(1, 4, 14)

#cols = ascii_uppercase
for row in range(0,len(trend_stats)):
    stats_sheet.add_sparkline('E{}'.format(row+2), {'range': 'Trends!{}{}:{}{}'.format(get_col(3),row+2, get_col(len(trends.columns)), row+2)})
    stats_sheet.set_row(row+2, 25)
    
writer.save()


# In[ ]:





# In[ ]:





# In[ ]:




