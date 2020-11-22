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


def get_rid_of_commas(column):
    
    rid_comma_mask = column.astype(str).str.contains(',',na = False)
    comma_mask_tuple = dfObj.loc[rid_comma_mask, column]
    comma_mask_index = list(comma_mask_tuple.index.values)
    comma_mask_list = []
    for i in comma_mask_tuple:
              i = i.replace(',','')
              
              comma_mask_list.append(i)
          
            
    commadf = pd.DataFrame(comma_mask_index)
    commadf.set_index(0,inplace=True)
    commadf['YTD_This_Year_Qty'] = comma_mask_list
    dfObj.update(commadf) #Fix all this before running






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
    lenlistofstores = len(listofstores)
    
    
    # layout the Window
    layout = [[sg.Text('Fixing columns and misread fields....')],
          [sg.ProgressBar(lenlistofstores, orientation='h', size=(20, 20), key='progbar')],
          [sg.Cancel()]]

# create the Window
    window = sg.Window('Custom Progress Meter', layout)
# loop that would normally do something useful   

    
    
    for stores in listofstores:
      #sg.OneLineProgressMeter('Reading pdfs and checking that data is in correct fields', stores+1, lenlistofstores, key='-IMAGE-', orientation='h')
      tempdf = dfObj[firstnum:secondnumber]
      tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
      firstnum = secondnumber
      secondnumber = indexofstores[counter]
      counter = counter + 1

      #counter.append(storeval)
       # check to see if the cancel button was clicked and exit loop if clicked
      event, values = window.read(timeout=0)
      if event == 'Cancel' or event == sg.WIN_CLOSED:
          break
        # update bar with loop value +1 so that bar eventually reaches the maximum
      window['progbar'].update_bar(counter + 1)
      
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
          
      try:#Need to fix the "Code" Column
             codemask = tempdf['Code'].astype(str).str.contains('#|NS|PV|Pure', na=False)
             codesplit = tempdf.loc[codemask,'Code'].str.split(" ", 1 ,expand = True)
             codesplit.columns= ['Code','Description']
             txt = "apple#banana#cherry#orange"
             
             x = txt.split("#", 1)
             tempdf.update(codesplit)





             dfObj.update(tempdf)
      except: print("didn't fix code column")
      
      
         
      finally:
          

          dfObj.update(tempdf)



          
          
    
      
    dfObj.update(tempdf)
    dfObj.dropna(subset = ["Code"], inplace = True)
    dfObj  = dfObj[dfObj.Code != 'Code']
              
 #Clean up the code column
    dfObj = dfObj[dfObj.Store_Name.str.startswith(r"Total", na = False)]
    #dfObj = dfObj[dfObj.Code.str.startswith(r"1|2|3|4|5|6|7|8|9|0",na=False)]
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Category",na=False)]
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Customer",na=False)]
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Code",na=False)]
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Total",na=False)]



    dfObj = dfObj.reset_index(drop = True)
    
    
    try:#From here on down fix before running
          comma_mask = dfObj['YTD_This_Year_Qty'].astype(str).str.contains(',',na = False)
          comma_mask_tuple = dfObj.loc[comma_mask, 'YTD_This_Year_Qty']
          comma_mask_index = list(comma_mask_tuple.index.values)
          comma_mask_list = []
          for i in comma_mask_tuple:
              i = i.replace(',','')
              
              comma_mask_list.append(i)
          
            
          commadf = pd.DataFrame(comma_mask_index)
          commadf.set_index(0,inplace=True)
          commadf['YTD_This_Year_Qty'] = comma_mask_list
            
          dfObj.update(commadf) #Fix all this before running
    
    except: print("uh oh!")
    try:
          comma_mask2 = dfObj['YTD_Last_Year_Qty'].astype(str).str.contains(',',na = False)
          comma_mask_tuple2 = dfObj.loc[comma_mask2, 'YTD_Last_Year_Qty']
          comma_mask_index2 = list(comma_mask_tuple2.index.values)
          comma_mask_list2 = []
          for i in comma_mask_tuple2:
              i = i.replace(',','')
              comma_mask_list2.append(i)
          
            
          commadf2 = pd.DataFrame(comma_mask_index2)
          commadf2.set_index(0,inplace=True)
          commadf2['YTD_Last_Year_Qty'] = comma_mask_list2
            
          dfObj.update(commadf2)
    except:print("uh oh!")
    try: #get_rid_of_commas(dfObj['FY_2020_Qty'])
          comma_mask3 = dfObj['FY_2020_Qty'].astype(str).str.contains(',',na = False)
          comma_mask_tuple3 = dfObj.loc[comma_mask3, 'FY_2020_Qty']
          comma_mask_index3 = list(comma_mask_tuple3.index.values)
          comma_mask_list3 = []
          for i in comma_mask_tuple3:
              i = i.replace(',','')
              
              comma_mask_list3.append(i)
          
            
          commadf3 = pd.DataFrame(comma_mask_index3)
          commadf3.set_index(0,inplace=True)
          commadf3['FY_2020_Qty'] = comma_mask_list3
            
          dfObj.update(commadf3)
    except:print("catestrophic failure")
    try: #get_rid_of_commas(dfObj['FY_2019_Qty'])
          comma_mask4 = dfObj['FY_2019_Qty'].astype(str).str.contains(',',na = False)
          comma_mask_tuple4 = dfObj.loc[comma_mask4, 'FY_2019_Qty']
          comma_mask_index4 = list(comma_mask_tuple4.index.values)
          comma_mask_list4 = []
          for i in comma_mask_tuple4:
              i = i.replace(',','')
              
              comma_mask_list4.append(i)
          
            
          commadf4 = pd.DataFrame(comma_mask_index4)
          commadf4.set_index(0,inplace=True)
          commadf4['FY_2019_Qty'] = comma_mask_list4
            
          dfObj.update(commadf4)
    except:print("catestrophic failure")
    
    try:     dfObj[['FY_2020_Qty','FY_2019_Qty', 'YTD_This_Year_Qty', 'YTD_Last_Year_Qty']] =dfObj[['FY_2020_Qty','FY_2019_Qty', 'YTD_This_Year_Qty', 'YTD_Last_Year_Qty']].apply(pd.to_numeric)
    except:print('Yikes!')
    finally:print('yay!')


    
    
    
    
    
    
    
    
    #set column types
    
  
    
    
   #Need to remove all the commas
    #ValueError: Unable to parse string "3,275" at position 124
    
    
    
    #workbook = xlsxwriter.Workbook(out_path)
   
    
    #worksheet = workbook.add_worksheet('All_Stores')
    #worksheet.write(dfObj,[cell_format])
    #workbook.close()



    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    dfObj.to_excel(writer, sheet_name='All Stores', index = False)
    
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['All Stores']

    # Add some cell formats.
    #currency_format = workbook.add_format({'num_format': '$#,##0'})
    #cell_format = workbook.add_format()
    #cell_format.set_align('left')    

    #cell_format = workbook.add_format({'align': 'left'})
    
    #per_format  =  workbook.add_format({'num_format': '0%'})
    
   # worksheet.set_column('A:A',cell_format)#Code
   # worksheet.set_column('B:B',cell_format)#Description
   # worksheet.set_column('C:C',cell_format)#Qty
    #worksheet.set_column('D:D',cell_format)#Sales
   # worksheet.set_column('E:E',cell_format)#Qty
   # worksheet.set_column('F:F',currency_format)#Sales
   # worksheet.set_column('G:G',cell_format)#%Change
   # worksheet.set_column('H:H',cell_format)#Qty
   # worksheet.set_column('I:I',per_format)#Sales

    
    (last_row, last_col) = dfObj.shape
    
    column_settings = [{'header': 'Code', }, 
                       {'header': 'Description', }, 
                       {'header': 'FY_2020_Qty',}, 
                       {'header': 'FY_2020_Sales',}, 
                       {'header': 'FY_2019_Qty', }, 
                       {'header': 'FY_2019_Sales', },
                       {'header': 'Per_Chg_Periods',},
                       {'header': 'YTD_This_Year_Qty'}, 
                       {'header': 'YTD_This_Yr_Sales',}, 
                       {'header': 'YTD_Last_Year_Qty',}, 
                       {'header': 'YTD_Last_Yr_Sales', }, 
                       {'header': 'Per_Chg_Yrs', }, 
                       {'header': 'Period', }, 
                       {'header': 'Store_Name', }]

    # Create a list of column headers, to use in add_table().
    #column_settings = [{'header': column} for column in df.columns]
     #Align cells left justified
    #format = workbook.add_format()
    #format.set_align('left')

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, last_row, last_col-1,{'columns': column_settings, 'style':'Table Style Light 11' }) 
   
    #worksheet.set_column(0, last_col - 1, format)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    # done with loop... need to destroy the window as it's still open
    window.close()
    
        #Popup that tells our users where the files are at
    sg.popup('View results at ' +  out_path)
      

          

#df = pd.DataFrame({'Numbers':    [1010, 2020, 3030, 2020, 1515, 3030, 4545],
#                   'Percentage': [.1,   .2,   .33,  .25,  .5,   .75,  .45 ],
#})

# Create a Pandas Excel writer using XlsxWriter as the engine.
#writer2 = pd.ExcelWriter(out_path, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
#df.to_excel(writer2, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
#workbook  = writer2.book
#worksheet = writer2.sheets['Sheet1']

# Add some cell formats.
#format1 = workbook.add_format({'num_format': '#,##0.00'})
#format2 = workbook.add_format({'num_format': '0%'})

# Note: It isn't possible to format any cells that already have a format such
# as the index or headers or any cells that contain dates or datetimes.

# Set the column width and format.
#worksheet.set_column('B:B', 18, format1)

# Set the format but not the column width.
#worksheet.set_column('C:C', None, format2)
#writer2.save()

    