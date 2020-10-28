#------------------------------------------#
# Title: Report Data Cleaning and Export
#        
# Desc: A script used to clean up PDFs
#       and export to csv.
# Change Log: (Who, When, What)
# DO, 
#------------------------------------------#

# -- PREP -- #
#pip install tabula-py in cmd prompt

import tabula
import re

import pandas as pd
import numpy as np

import os


folder = 'E:/MSBA_UW/Project/PDFs/'
dfObj = pd.DataFrame()
pathvalue = 0
pathcounter = []
paths = [folder + fn for fn in os.listdir(folder) if fn.endswith('.pdf')]


for path in paths:
    df = tabula.io.read_pdf(path, pages = 'all', pandas_options={'header': None})
    totalelementsdf = len(df)
    counter = []

    value = 0

  # use while loop to add the list of dataframes together
#dfObj = pd.DataFrame(df[value])
    maxnum = (len(df)-1)
    while value <= maxnum:   
                dfObj = dfObj.append(pd.DataFrame(df[value]), ignore_index=True)
                value += 1
                counter.append(value)
    #dfObj = dfObj.append(pd.DataFrame(df[pathvalue]),ignore_index=True)
    #pathvalue += 1
#    pathcounter.append(pathvalue)
           

# -- DATA -- #

#df = tabula.io.read_pdf(r"E:\MSBA_UW\Project\PDFs\Period 1 2020 SBA Dietz 123.pdf",pages = 'all', pandas_options={'header': None})



 #rename the column names to correct columns
dfObj.columns = ['Code', 'Description', 'FY_2020_Qty', 'FY_2020_Sales', 'FY_2019_Qty', 'FY_2019_Sales', 'Per_Chg_Periods', 'YTD_This_Year_Qty', 'YTD_This_Yr_Sales', 'YTD_Last_Year_Qty', 'YTD_Last_Yr_Sales', 'Per_Chg_Yrs']


#Find the period and append to a period column
listofperiods = []
listofperiods = dfObj[dfObj['FY_2020_Qty'].astype(str).str.contains(r'Period', na = False)]
listofperiods = listofperiods['FY_2020_Qty']
indexofperiods = list(listofperiods.index.values)

#add a column for store (and period)
dataframelength = len(dfObj)
storecollist = []
periodcollist = []
for i in range(dataframelength):
    storecollist.append(i)
    periodcollist.append(i)
storecol = pd.DataFrame(storecollist)
periodcol = pd.DataFrame(periodcollist)
dfObj['Period'] = periodcol
dfObj['Store_Name'] = storecol





periodval = 0
periodfirstnum = 1
periodsecondnum = indexofperiods[periodval]
numofperiods = len(listofperiods)
while periodval <= numofperiods:
    perioddf = dfObj[periodfirstnum:periodsecondnum]
    perioddf = perioddf.assign(Period = listofperiods[periodfirstnum])
    periodfirstnum = periodsecondnum
    periodval += 1
    counter.append(periodval)
    try:
        periodsecondnum = indexofperiods[periodval]
        dfObj.update(perioddf)
    except:
        dfObj.update(perioddf)

#Make sure that store names are in the Description column notice there is a space after Total
mask = dfObj['Code'].str.startswith(r'Total ', na=False)
dfObj.loc[mask,'Description'] = dfObj['Code']


listofstores =[]#'Total PET CLUB WAREHOUSE', 'Total RIO GRANDE SERVICE CENTER', 'Total SUNBURST PET SUPPLIES'
#find all the store names, this will tell us how to find the beginning and end of each dataframe, as well as populate our last column
listofstores = dfObj[dfObj['Description'].str.startswith(r"Total", na = False)]
listofstores=listofstores['Description']
indexofstores = list(listofstores.index.values)
storeval =0    
firstnum = 0
secondnumber = indexofstores[storeval] 
numofstores = len(listofstores)
while storeval < numofstores:

      tempdf = dfObj[firstnum:secondnumber]
      tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
      firstnum = secondnumber
      storeval += 1
      counter.append(storeval)
      
      
      #desincode_mask = tempdf['Code'].astype(str).str.contains(r' ', regex = False, na = False)
      #codesplit = tempdf.loc[desincode_mask, 'Code'].str.split(' ', 1, expand=True)
      #codesplit.columns = ['Code','Description']
      #tempdf.update(codesplit)      

            #Fix columns right to left
      try:
          mask = tempdf['YTD_Last_Year_Qty'].astype(str).str.contains(r'%', na=False) 
          tempdf.loc[mask, 'Per_Chg_Yrs'] = tempdf['YTD_Last_Year_Qty']

          mask2 = tempdf['YTD_This_Year_Qty'].astype(str).str.startswith(r'$', na=False)
          tempdf.loc[mask2, 'YTD_Last_Yr_Sales'] = tempdf['YTD_This_Year_Qty']
     
      #The issue here is that the next page is getting being overwritten when tempdf['whatever'] = tempsplit[0]
          mask3 = tempdf['Per_Chg_Periods'].astype(str).str.startswith(r'$', na = False)
          tempsplit = tempdf.loc[mask3,'Per_Chg_Periods'].str.split(expand=True)
      finally:
          print('string contains successful')
      try:#Needs to be two columns otherwise will not work
              tempsplit.columns = ['YTD_This_Yr_Sales','YTD_Last_Year_Qty']
              tempdf.update(tempsplit)
      except:
              print('error: could not split into two columns')
      finally:
          print('successfully updated YTD_This_Yr_Sales and YTD_Last_Year_Qty')
      tempdf['Per_Chg_Periods'] = tempdf['Per_Chg_Periods'].astype(str)
      mask35 = tempdf['Per_Chg_Periods'].astype(str).str.isdigit()
      tempdf.loc[mask35,'YTD_Last_Year_Qty'] = tempdf['Per_Chg_Periods']
      
      
      mask4 = tempdf['FY_2019_Sales'].astype(str).str.contains(r' ',regex = False, na= False)

      tempsplit2 = tempdf.loc[mask4,'FY_2019_Sales'].str.split(expand=True)
      try:#Needs to be two columns otherwise will not work
              
              tempsplit2.columns = ['YTD_This_Year_Qty','YTD_This_Yr_Sales']
              tempdf.update(tempsplit2)
      except:
              print('error: could not split into two columns')
      finally:
          print('successfully updated')
      try: 
          tempdf.dropna(subset = ["FY_2019_Sales"], inplace = True)

          mask45 = tempdf['FY_2019_Sales'].astype(str).str.isdigit()
          tempdf.loc[mask45,'YTD_This_Year_Qty'] = tempdf['FY_2019_Sales']
          tempdf['FY_2019_Qty'].fillna('0', inplace = True)
          mask6 = tempdf['FY_2019_Qty'].astype(str).str.contains(r'%',regex= False, na= False)
          tempdf.loc[mask6, 'Per_Chg_Periods'] = tempdf['FY_2019_Qty']

      finally:

          
      
          mask7 = tempdf['FY_2020_Sales'].astype(str).str.contains(r' ',regex= False, na= False)
    
          tempsplit4 = tempdf.loc[mask7,'FY_2020_Sales'].str.split(expand=True)
      try:#Needs to be two columns otherwise will not work
              tempsplit4.columns = ['FY_2019_Qty','FY_2019_Sales']
              tempdf.update(tempsplit4)
             
      except:
              print('error: could not split into two columns')
      finally:
          print('successfully updated')
      
      mask8 = tempdf['FY_2020_Qty'].astype(str).str.contains(r' ',regex= False, na= False)
    
      tempsplit4 = tempdf.loc[mask7,'FY_2020_Qty'].str.split(expand=True)
      try:#Needs to be two columns otherwise will not work
              tempsplit4.columns = ['FY_2020_Qty','FY_2020_Sales']
              tempdf.update(tempsplit4)
      except:
          print('error: could not split into two columns')
      finally:
          print('successfully updated')

      try:
          secondnumber = indexofstores[storeval]
          dfObj.update(tempdf)
      except:
          print("continue on with life")    
          dfObj.update(tempdf)
          dfObj.dropna(subset = ["Code"], inplace = True)
          dfObj  = dfObj[dfObj.Code != 'Code']
          dfObj = dfObj.reset_index(drop = True)
          
          #Run this because I'm lazy and want a clean dataset
          dfObj = dfObj[dfObj.Store_Name.str.startswith(r"Total", na = False)]
          
          
          
          out_path = r"E:/MSBA_UW/Project/temp-excel.xlsx"
          writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
          dfObj.to_excel(writer, sheet_name='All Stores')
          writer.save()
      

          

          











