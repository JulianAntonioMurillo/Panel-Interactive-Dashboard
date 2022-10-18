#!/usr/bin/env python
# coding: utf-8

# In[2]:


# Analysis of THD Online Sales by Online SKU
### For Brian/BOBJ
### Load in Correct Data and just reorganize it
import sys
sys.path.append(r'C:\Users\jmurillo\AppData\Local\Programs\Python\Python310\Lib\site-packages')


import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sn
from bs4 import BeautifulSoup
import dataframe_image as dfi
from plotnine import *
import win32com.client as win32
# from sklearn.linear_model import LinearRegression # Linear Regression Model
# from sklearn.preprocessing import StandardScaler #Z-score variables
# from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score #model evaluation
import re
from matplotlib.pyplot import figure
# import altair as alt
# import datum
import hvplot.pandas
# from vega_datasets import data as vds
import panel as pn
pn.extension('tabulator')
import requests
import xlrd
import io

# import dash
# import dash_core_components as dcc
# import dash_html_components as html
# from dash.dependencies import Output,Input


# In[44]:


url = "https://raw.githubusercontent.com/JulianAntonioMurillo/Panel-Interactive-Dashboard/main/internet_metrics_df.csv"
internet_metrics = pd.read_csv(url)
internet_metrics.head()


# In[45]:


# Reformat

# internet_metrics = pd.read_excel(r"C:\Users\jmurillo\Desktop\Misc\THD_Internet_Metrics_mktg\IndividualFiles\InternetRatingsReviews.xlsx")
# internet_metrics.head(8)








# Change Dtypes/replace NA's with 0
#internet_metrics.update(internet_metrics[['Online Count of 1 Star Reviews +','Online Count of 2 Star Reviews +','Online Count of 3 Star Reviews +',
#                                          'Online Count of 4 Star Reviews +','Online Count of 5 Star Reviews +','Online PIP Visits +',
#                                          'Online PIP Conversion Rate +','Online Product Interaction Visits +','Online Gross Demand $ +',
#                                          'Online Order Units +','Online Sales $ +','Online Cancel Units +','Online Cancel $ +',
#                                          'Online Return $ +','Online Return Units +', 'is_multi_count','Week','Week_1','Year_1']].fillna(0))

#internet_metrics = internet_metrics.astype({'Online Count of 1 Star Reviews +':'int64','Online Count of 2 Star Reviews +':'int64','Online Count of 3 Star Reviews +':'int64',
#                                          'Online Count of 4 Star Reviews +':'int64','Online Count of 5 Star Reviews +':'int64','Online PIP Visits +':'int64',
#                                          'Online PIP Conversion Rate +':'float64','Online Product Interaction Visits +':'int64','Online Gross Demand $ +':'float64',
#                                          'Online Order Units +':'int64','Online Sales $ +':'float64','Online Cancel Units +':'int64','Online Cancel $ +':'float64',
#                                          'Online Return $ +':'float64','Online Return Units +':'int64','Online Avg Rating  +':'float64',
#                                           'Online Current List Price $ +':'float64','is_multi_count':'int64'})

# Convert Date to DateTime
internet_metrics['DateTime_Start']= pd.to_datetime(internet_metrics['DateTime_Start'])
internet_metrics['DateTime_End']= pd.to_datetime(internet_metrics['DateTime_End'])
#Reformat month year to month year
internet_metrics['Month_Year']= pd.to_datetime(internet_metrics['Month_Year'], errors = 'coerce')
internet_metrics['Month_Year'] = internet_metrics['Month_Year'].dt.strftime('%m/%Y')
internet_metrics['Month_Year']= pd.to_datetime(internet_metrics['Month_Year'])

internet_metrics.columns = internet_metrics.columns.str.replace(' ','_')


# In[46]:


# Find top 10 selling SKU's
group_by_sum = pd.DataFrame(internet_metrics.groupby(['Online_THD_SKU+'])['Online_Sales_$_+'].sum())
group_by_sum = group_by_sum.reset_index()
group_by_sum = group_by_sum.sort_values(by=['Online_Sales_$_+'], ascending=False)
group_by_sum.head(10)


# In[47]:


# Select top 10 online selling items
top_10 = internet_metrics[(internet_metrics['Online_THD_SKU+'].str.contains("431429-SG APC 128OZ")==True) | 
                       (internet_metrics['Online_THD_SKU+'].str.contains("883387-SIMPLE GREEN APC 320OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1002075713-SMPL GRN OUTDR ODOR ELIMINATOR 128OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("853534-SG PRO HEAVY DUTY 128OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("854029-SG PRO3PLUS ANTIBAC&DISINFECT 128OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("435909-SG APC CONCEN SPY 32OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1000017290-SG APC 640OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1002332519-5 GAL. EXTREME AIRCRAFT AND PRECISIO")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1002075704-SMPL GRN OUTDR ODOR ELIMINATOR 32OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1001700777-1 GAL. CONCENTRATED ALL-PURPOSE CLEA")==True)]

# Select top 5 online selling items
top_5 = internet_metrics[(internet_metrics['Online_THD_SKU+'].str.contains("431429-SG APC 128OZ")==True) | 
                       (internet_metrics['Online_THD_SKU+'].str.contains("883387-SIMPLE GREEN APC 320OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("1002075713-SMPL GRN OUTDR ODOR ELIMINATOR 128OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("853534-SG PRO HEAVY DUTY 128OZ")==True)|
                       (internet_metrics['Online_THD_SKU+'].str.contains("854029-SG PRO3PLUS ANTIBAC&DISINFECT 128OZ")==True)]
print(top_5.shape,top_10.shape)


# In[48]:


# plot top 10 over all time

#sn.set_palette("colorblind")
#ax = sn.lineplot(x = "DateTime_End", y = "Online_PIP_Visits_+", hue = "Online_THD_SKU+", data = top_10)
#plt.show()


# In[49]:


# Interactive chart with drop-down menu all time by week end date
all_time_visits = top_10.hvplot(x='DateTime_End', y='Online_PIP_Visits_+', groupby='Online_THD_SKU+', kind='line',
                               xlabel = 'Date (Weekly)', ylabel = 'Online PIP Visits').opts(bgcolor='lightgray', show_grid = True)
all_time_visits


# In[50]:


# Interactive chart with drop-down menu all time by average grouped by month
# Still top 10 selling items
month_avg = pd.DataFrame(top_10.groupby(['Month_Year','Online_THD_SKU+'])['Online_PIP_Visits_+'].mean())


month_avg = month_avg.reset_index()
month_avg.rename(columns={"Online_PIP_Visits_+": "Online_PIP_Visits_+_monthly_avg"}, inplace = True)

month_avg.tail()
month_avg = month_avg.sort_values(by=['Month_Year'], ascending=True)

month_avg_visits = month_avg.hvplot(x='Month_Year', y='Online_PIP_Visits_+_monthly_avg', groupby='Online_THD_SKU+', kind='line',
                                   xlabel = 'Date (Month)', ylabel = 'Average Online PIP Visits').opts(bgcolor='lightgray', show_grid = True)
month_avg_visits
# This plot gives me a better idea month to month what the goal would be for PIP visits for the top 10 SKU's
# you can also see the seasonality... the last half of the graph shows the later months for all years 2017-2022. They're significantly lower than the middle or first part of the graph
# If you visualize the graph as a loop... you can see that sales go down for outdoor eliminator during the fall/winter months... then pick back up in January into spring/summer.


# In[7]:


# Group by month and sum sales and canceled orders
#month_sum = pd.DataFrame(top_10.groupby(['Year_1'])[['Online_Sales_$_+','Online_Cancel_$_+']].sum())
#month_sum = month_sum.reset_index()
# Plot that shit
# multiple line plots

#plt.plot( 'Year_1', 'Online_Sales_$_+', data=month_sum, marker='', color='skyblue', linewidth=2, label = "Online Sales")
#plt.plot( 'Year_1', 'Online_Cancel_$_+', data=month_sum, marker='', color='red', linewidth=2, linestyle='dashed', label="Online Cancellations")

# show legend
#plt.legend()

# show graph
#plt.show()


# In[51]:


# Group by month and sum sales and canceled orders
ratings_df = pd.DataFrame(top_10[['Online_THD_SKU+','Online_Count_of_1_Star_Reviews_+',
       'Online_Count_of_2_Star_Reviews_+', 'Online_Count_of_3_Star_Reviews_+',
       'Online_Count_of_4_Star_Reviews_+', 'Online_Count_of_5_Star_Reviews_+']])


ratings_df_sum = pd.DataFrame(ratings_df.groupby(['Online_THD_SKU+'])[['Online_Count_of_1_Star_Reviews_+',
       'Online_Count_of_2_Star_Reviews_+', 'Online_Count_of_3_Star_Reviews_+',
       'Online_Count_of_4_Star_Reviews_+', 'Online_Count_of_5_Star_Reviews_+']].sum())
ratings_df_sum = ratings_df_sum.reset_index()
ratings_df_sum.columns = ratings_df_sum.columns.str.replace('+','')
ratings_df_sum.rename(columns={"Online_Count_of_1_Star_Reviews_": "1_Star",
                              "Online_Count_of_2_Star_Reviews_":"2_Star",
                              "Online_Count_of_3_Star_Reviews_":"3_Star","Online_Count_of_4_Star_Reviews_":"4_Star",
                               "Online_Count_of_5_Star_Reviews_":"5_Star"}, inplace = True)

# Calculate percentages for each out of the row total
cols = ['1_Star','2_Star','3_Star','4_Star', '5_Star']
ratings_df_sum[cols] = ratings_df_sum[cols].div(ratings_df_sum[cols].sum(axis=1), axis=0).multiply(100)
ratings_df_sum.head(12)


# In[53]:


# Create visual for total count of ratings across all rating types over time
col_names = ['Online_Count_of_1_Star_Reviews_+',
       'Online_Count_of_2_Star_Reviews_+', 'Online_Count_of_3_Star_Reviews_+',
       'Online_Count_of_4_Star_Reviews_+', 'Online_Count_of_5_Star_Reviews_+']
top_10['number_of_ratings']= top_10[col_names].sum(axis=1)
# Plot data:
# Interactive chart with drop-down menu all time by MONTH end date

month_rat_sum = pd.DataFrame(top_10.groupby(['Month_Year','Online_THD_SKU+'])['number_of_ratings'].sum())


month_rat_sum = month_rat_sum.reset_index()
month_rat_sum.rename(columns={"number_of_ratings": "number_of_ratings_monthly_sum"}, inplace = True)

month_rat_sum.tail()
month_rat_sum = month_rat_sum.sort_values(by=['Month_Year'], ascending=True)

month_rat_sum_plot = month_rat_sum.hvplot(x='Month_Year', y='number_of_ratings_monthly_sum', groupby='Online_THD_SKU+', kind='line', color = 'green', 
                                          xlabel = 'Date (Month)', ylabel = 'Count of Ratings').opts(bgcolor='lightgray', show_grid = True)
month_rat_sum_plot


# all_time_ratings_count = top_10.hvplot(x='DateTime_End', y='number_of_ratings', groupby='Online_THD_SKU+', kind='line')
# all_time_ratings_count


# In[54]:


# Pivot Table
from pivottablejs import pivot_ui
import ipypivot as pt


# In[55]:


pivot_table_top_10 = pivot_ui(top_10)
pivot_table_top_10


# ### ------------------------------------------- Save Interactive Charts and Create Dashboard -------------------------------------------

# In[57]:


# Save interactive chart as HTML for functionality
#### from bokeh.resources import INLINE
#### hvplot.save(zoink_plot, 'OnlinePIP_Visits_Top10_SKUs_THD.html', resources=INLINE)
os.chdir(r"C:\Users\jmurillo\Desktop\Misc\THD_Internet_Metrics_mktg\IndividualFiles")
#Layout using Template
template = pn.template.FastListTemplate(
    title='Internet Metrics - The Home Depot', 
    sidebar=[pn.pane.Markdown("# The Home Depot Online Sales"), 
             pn.pane.Markdown("#### This Dashboard displays summary statistics from Home Depot's Online Sales Performance. The following is data pulled directly from HomeDepot/Vendor Drill, something we were previously lacking insight on. We have great data on how we sell our products to them, but now how our products perform in their stores/on HomeDepot.com"), 
             pn.pane.PNG('Logo_SimpleGreen_White_Outline-Green.png', sizing_mode='scale_both'),
             pn.pane.PNG('TheHomeDepot.png', sizing_mode='scale_both')],
    main=[pn.Row(pn.Column(month_avg_visits, all_time_visits)),
          pn.Row(pn.Column(month_rat_sum_plot, ratings_df_sum))],
    accent_base_color="#88d8b0",
    header_background="#88d8b0",
)
# template.show()
template.servable();


# In[20]:


os.chdir(r"C:\Users\jmurillo\Desktop\HomeDepot_DashBoard\dash_env\IndividualFiles")


# In[21]:


# Save individual interactive plots 
from bokeh.resources import INLINE
hvplot.save(month_rat_sum_plot, 'month_rat_sum_plot.html', resources=INLINE)


# In[ ]:




