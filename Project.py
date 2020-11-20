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
from datetime import date

import os
import PySimpleGUI as sg
import xlsxwriter

sg.theme('Purple')   # Add a little color to your windows


today = str(date.today())
layout = [      
#    [sg.Menu(menu_def, tearoff=True)],      
    [sg.Text('PDF to Excel Converter', size=(30, 1), justification='center', font=("Helvetica", 25), relief=sg.RELIEF_RIDGE)],    
    [sg.Text('SPA report option outputs to excel file with data cleaned and stores added')],      
    [sg.Text('Quick scan is a straight pdf to excel conversion')],       
    [sg.Frame(layout=[    
      #Radio buttons need to be the same group ID otherwise a user will be able to select both of them. "GROUP_ID" is the group_id 
    [sg.Radio('SPA Report', "GROUP_ID", default=True, size=(11,1)), sg.Radio('Quick scan', "GROUP_ID")]], title='Options',title_color='black', relief=sg.RELIEF_SUNKEN, 
        tooltip=' Complete scan may take more than 60 minutes if records exceed 100 pages ')], 
    [sg.Checkbox('output store performance sheet', size=(70,1))],   
    [sg.Checkbox('File to convert from is Excel', size=(70,1))],     
    [sg.Text((''), size=(25, 2), text_color= 'red'),      
       ],      
    [sg.Text('_'  * 80)],      
    [sg.Text('Choose a name for your report (lastname_report)')],     
    [sg.Input(today+'-SPA-report')],
    [sg.Text('Choose a location to save your results', size=(35, 1))],      
    [sg.Text('Your Folder', size=(15, 1), auto_size_text=False, justification='right'),      
        sg.InputText(''), sg.FolderBrowse()],
    [sg.Text('_'  * 80)],      
    [sg.Text('Choose a folder for where pdfs are stored', size=(35, 1))],      
    [sg.Text('Your Folder', size=(15, 1), auto_size_text=False, justification='right'),      
        sg.InputText(''), sg.FolderBrowse()],      
    [sg.Submit(tooltip='Click to submit this window'), sg.Cancel()]    
]      


window = sg.Window('PDF Converter', layout, default_element_size=(40, 1), grab_anywhere=False,keep_on_top= True)      

event, values = window.read()    
  
    #Definitions of the dict values variables
    #value[0] is the radio button for complete scan. Either True/False 
    #value[1] is the radio button for quick scan. Either True/False
    #value[2] the check box for the performance reports Etiher True/False
    #value[3] the check box for if the origional file is an excel Etiher True/False
    #value[4] is the file naming field
    #value[5] is the location selected for results to be saved
    #value[6] is the location selected for where the records are located
    
window.close()

folder = values[6] + '/'
#Add error handling if user cancels
if not folder:
    sg.popup_cancel("Cancelled: Must browse to a folder with pdfs")

out_path = values[5] + '/' + values[4] + '.xlsx'
if not out_path:
    sg.popup_cancel("Cancelled: Must select a location to output excel file")


#Excel reading
#if values[3] is True:
#    paths = [folder + fn for fn in os.listdir(folder) if fn.endswith('.xlsx')]
#    for path in paths:
#
#        dfObj = pd.read_excel(paths)









dfObj = pd.DataFrame()
pathvalue = 0
pathcounter = []
paths = [folder + fn for fn in os.listdir(folder) if fn.endswith('.pdf')]
if not paths:
    sg.popup_cancel("Cancelled: Must browse to a folder with pdfs")



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
                sg.OneLineProgressMeter('Processing Reports', value + 1, maxnum, 'key', orientation = 'h')

                
                
                
                
if values[1] is True:
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
    dfObj.to_excel(writer, sheet_name='All Stores')
    writer.save()
    sg.popup('Completed!', 'Converted %s pages'%(maxnum))


if values[0] is True:
    dfObj.columns = ['Code', 'Description', 'FY_2020_Qty', 'FY_2020_Sales', 'FY_2019_Qty', 'FY_2019_Sales', 'Per_Chg_Periods', 'YTD_This_Year_Qty', 'YTD_This_Yr_Sales', 'YTD_Last_Year_Qty', 'YTD_Last_Yr_Sales', 'Per_Chg_Yrs']
    #Find the period and append to a period column
    listofperiods = []
    listofperiods = dfObj[dfObj['FY_2020_Qty'].astype(str).str.contains(r'Period', na = False)].copy()
    listofperiods = listofperiods['FY_2020_Qty'].copy()

    indexofperiods = list(listofperiods.index.values)

    
    #add a column for store (and period)
    dataframelength = len(dfObj)
    #Add a final period and index because otherwise it won't fill in a period to the end of the dataframe
    indexofperiods.append(dataframelength)
    #Get the last period used in the dataframe
    lastperiod = listofperiods.iat[-1]
    lastperiod = lastperiod[1:-1]
    #Combine this with the last index value and aadd to listofperiods
    listofperiods.loc[dataframelength] = lastperiod
    
    storecollist = []
    periodcollist = []
    for i in range(dataframelength):
        storecollist.append(i)
        periodcollist.append(i)
    storecol = pd.DataFrame(storecollist)
    periodcol = pd.DataFrame(periodcollist)
    dfObj['Period'] = periodcol
    dfObj['Store_Name'] = storecol




#Assign periods
    periodval = 0
    periodfirstnum = indexofperiods[0]
    periodsecondnum = indexofperiods[periodval]
    numofperiods = len(listofperiods)
    while periodval <= numofperiods:
            sg.OneLineProgressMeter('Determining periods and adding a column', periodval+1, numofperiods, key='-IMAGE-', orientation='h')
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
    secondnumber = (indexofstores[storeval]) 
    numofstores = len(listofstores)

    #while storeval < numofstores:

#      tempdf = dfObj[firstnum:secondnumber]
#      tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
#     firstnum = secondnumber
#     storeval += 1
#     secondnumber = indexofstores[storeval]

    counter = 0
    #for stores in listofstores: #Did I double name a variable

    for stores in listofstores:
      #sg.OneLineProgressMeter('Reading pdfs and checking that data is in correct fields', stores+1, len(listofstores), key='-IMAGE-', orientation='h')
      tempdf = dfObj[firstnum:secondnumber]
      tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
      firstnum = secondnumber
      secondnumber = indexofstores[counter]
      counter = counter + 1

      #counter.append(storeval)
      
      
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
          dfObj.update(tempdf)
          print("continue on with life")
          
      #try:#Need to fix the "Code" Column
             # mask = dfObj['Code'].str.startswith(r'Total ', na=False)

      finally:
          dfObj.update(tempdf)
    
      
    dfObj.update(tempdf)
    dfObj.dropna(subset = ["Code"], inplace = True)
    dfObj  = dfObj[dfObj.Code != 'Code']
    dfObj = dfObj.reset_index(drop = True)
              
 #Run this because I'm lazy and want a clean dataset
    dfObj = dfObj[dfObj.Store_Name.str.startswith(r"Total", na = False)]
         
    
    
    
    
    
    #workbook = xlsxwriter.Workbook(out_path)
   
    
    #worksheet = workbook.add_worksheet('All_Stores')
    #worksheet.write(dfObj,[cell_format])
    #workbook.close()


    
    
    writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
    
    dfObj.to_excel(writer, sheet_name='All Stores', index = False)
    
    workbook = writer.book
    worksheet = writer.sheets['All Stores']
    
    (last_row, last_col) = dfObj.shape
    
    column_settings = [{'header': column} for column in dfObj.columns]

    # Create a list of column headers, to use in add_table().
    #column_settings = [{'header': column} for column in df.columns]
     #Align cells left justified
    #format = workbook.add_format()
    #format.set_align('left')

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, last_row, last_col-1,{'columns': column_settings})
   
    #worksheet.set_column(0, last_col - 1, format)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    
    
        #Popup that tells our users where the files are at
    sg.popup('View results at ' +  out_path)
      

          

          











