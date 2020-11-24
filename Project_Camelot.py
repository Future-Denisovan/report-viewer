
import os
import PySimpleGUI as sg
import xlsxwriter
import camelot
import pandas as pd
from datetime import date
import pandas as pd 



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
    [sg.Checkbox('Output store performance sheet', size=(70,1))],   
    [sg.Checkbox('File to convert from is Excel', size=(70,1))],     
    [sg.Text('Excel File', size=(8, 1)), sg.Input(), sg.FileBrowse()],
    [sg.Text('Number of Sheets', size=(13, 1)),sg.InputText('1', size= (5,1))],      
    [sg.Text((''), size=(25, 1), text_color= 'red'),      
       ],      
    [sg.Text('Choose a name for your report:')],     
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
    #value[4] is the excel file, if any is selected
    #value[5] is the number of sheets in the excel file
    #value[6] is the file naming field
    #value[7] is the location selected for results to be saved
    #value[8] is the location selected for where the records are located
    
window.close()

folder = values[8] + '/'
out_path = values[7] + '/' + values[6] + '.xlsx'
#Add error handling if user cancels
if not folder:
    sg.popup_cancel("Cancelled, user exited")
if not out_path:
    sg.popup_cancel("Cancelled: Must select a location to output excel file")

##################################DEF
def reading_excel_SPA_report(values):
    print('reading excel spa report')

    #create a list of values for number of sheets in excel file   
    #sheetlist=[]
    numberofsheets = int(values[5])
    #l = -1
    #while l < numberofsheets:
    #    l=l+1
    #    print(l)
    #    sheetlist.append(l)
    #sheetset = set(sheetlist)
    #sheetdict = dict.fromkeys(sheetset)
    pathforexcel=values[4]
    dfObj = pd.read_excel(pathforexcel, header = 9, skiprows=(0))
    
    return(dfObj)
##################################DEF
def quick_scan(dfObj):
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
    dfObj.to_excel(writer, sheet_name='All Stores')
    writer.save()
    sg.popup('Completed! See results at '+ out_path)
##################################


bignum = 1000000
smallnum = 0
while smallnum < bignum:
      smallnum = smallnum + 1
      sg.popup_animated('E:\MSBA_UW\Project\Project_Folder\SpecialProject\Fetch.gif',message='Please wait while I fetch that...',time_between_frames=100,keep_on_top=True)
sg.popup_animated(image_source=None)

pathvalue = 0
pathcounter = []
if values[3] is False:
    paths = [folder + fn for fn in os.listdir(folder) if fn.endswith('.pdf')]
if not paths:
    sg.popup_cancel("Cancelled: Must browse to a folder with pdfs")


if values[3] is True:
    reading_excel_SPA_report(values)
    

if values[3] is False:
    dfObj = pd.DataFrame()
    for path in paths:
        sg.OneLineProgressMeter('Processing Reports', pathvalue + 1, len(paths), 'key', orientation = 'h',size=(70,4))#
        #sg.popup_animated('E:\MSBA_UW\Project\Project_Folder\SpecialProject\Fetch.gif',message='Please wait while I fetch that...',background_color='Purple',time_between_frames=100,keep_on_top=True)
        tables = camelot.read_pdf(path, pages = '1-end', flavor="stream" )#,strip_text=','
        tablecounter = 0
        
        listoftables = tables.n
        counter = []
        value = 0
        pathvalue = pathvalue + 1

        while tablecounter < listoftables-1:
            sg.OneLineProgressMeter('Processing Reports', tablecounter + 1, listoftables, 'key', orientation = 'h',size=(70,4))
            dfObj = dfObj.append(tables[tablecounter].df,ignore_index=True)
            tablecounter = tablecounter +1
            value += 1
            counter.append(value)

              
if values[1] is True:
    quick_scan(dfObj)
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
        sg.OneLineProgressMeter('Adding new columns... ', i + 1, dataframelength, 'key', orientation = 'h')

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
    #mask = dfObj['Code'].str.startswith(r'Total ', na=False)
    #dfObj.loc[mask,'Description'] = dfObj['Code']


    listofstores =[]#'Total PET CLUB WAREHOUSE', 'Total RIO GRANDE SERVICE CENTER', 'Total SUNBURST PET SUPPLIES'
    #find all the store names, this will tell us how to find the beginning and end of each dataframe, as well as populate our last column
    listofstores = dfObj[dfObj['Description'].str.startswith(r"Total", na = False)]
    listofstores=listofstores['Description']
    indexofstores = list(listofstores.index.values)
    storeval =0    
    firstnum = 0
    storecounter = []
    secondnumber = (indexofstores[storeval]) 
    numofstores = len(listofstores)
    

    
    
    
    
    counter = 0
    #for stores in listofstores: #Did I double name a variable
    lenlistofstores = len(listofstores)
    
    
    # layout the Window
    #layout = [[sg.Text('Fixing columns and misread fields....')],
    #      [sg.ProgressBar(lenlistofstores, orientation='h', size=(20, 20), key='progbar')],
    #      [sg.Cancel()]]

    # create the Window
    #window = sg.Window('Custom Progress Meter', layout)
    # loop that would normally do something useful   
    
    
    
    for stores in listofstores:
      #sg.OneLineProgressMeter('Reading pdfs and checking that data is in correct fields', stores+1, lenlistofstores, key='-IMAGE-', orientation='h')
      tempdf = dfObj[firstnum:secondnumber]
      tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
      firstnum = secondnumber
      secondnumber = indexofstores[counter]
      counter = counter + 1
      dfObj.update(tempdf)
      #event, values = window.read(timeout=0)
      #if event == 'Cancel' or event == sg.WIN_CLOSED:
          #break
        # update bar with loop value +1 so that bar eventually reaches the maximum
      #window['progbar'].update_bar(counter + 1)
#
    tempdf = dfObj[firstnum:secondnumber]
    tempdf = tempdf.assign(Store_Name=listofstores[secondnumber])
    dfObj.update(tempdf)

      #counter.append(storeval)
       # check to see if the cancel button was clicked and exit loop if clicked
      
      
      
      
    #dfObj.update(tempdf)
    #dfObj.dropna(subset = ["Code"], inplace = True) Does nothing
    dfObj  = dfObj[dfObj.Code != 'Code']
              
    #Clean up the code column
    #dfObj = dfObj[dfObj.Store_Name.str.startswith(r"Total", na = False)] Does nothing
    #dfObj = dfObj[~dfObj.Code.str.startswith(r"Category",na=False)] Does nothing
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Customer",na=False)] #removed one row
    dfObj = dfObj[~dfObj.Code.str.startswith(r"Code",na=False)]
    dfObj = dfObj[~dfObj.Code.str.startswith(r"T",na=False)]

    #Clean up the Fiscal Year and Periods from the FY2020Qty Col
    dfObj = dfObj[~dfObj.FY_2020_Qty.str.startswith(r"Fiscal", na = False)]
    dfObj = dfObj[~dfObj.FY_2020_Qty.str.startswith(r"Period",na=False)]
    dfObj = dfObj[~dfObj.FY_2020_Sales.str.startswith(r"T",na=False)]
    dfObj = dfObj[~dfObj.FY_2020_Sales.str.startswith(r"F",na=False)]
    dfObj = dfObj[~dfObj.FY_2020_Sales.str.startswith(r"P",na=False)]



    
    #Clean up Description and Fy2019Qty
    dfObj = dfObj[~dfObj.FY_2019_Qty.str.startswith(r"Tuffy", na = False)]
    dfObj = dfObj[~dfObj.Description.str.startswith(r"Category",na=False)]
    dfObj = dfObj[~dfObj.Description.str.startswith(r"Total",na=False)]
    dfObj = dfObj[~dfObj.FY_2019_Sales.str.startswith(r"P",na=False)]
    dfObj = dfObj[~dfObj.FY_2019_Sales.str.startswith(r"T",na=False)]
    dfObj = dfObj[~dfObj.Description.str.startswith(r"USE",na=False)]
    dfObj = dfObj[~dfObj.FY_2019_Qty.str.startswith(r"P",na=False)]
    dfObj = dfObj[~dfObj.FY_2019_Qty.str.startswith(r"T",na=False)]

    dfObj = dfObj.reset_index(drop = True)
    
    
    try:     dfObj[['FY_2020_Qty','FY_2019_Qty', 'YTD_This_Year_Qty', 'YTD_Last_Year_Qty']] =dfObj[['FY_2020_Qty','FY_2019_Qty', 'YTD_This_Year_Qty', 'YTD_Last_Year_Qty']].apply(pd.to_numeric)
    except:print('Yikes!')


 # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    dfObj.to_excel(writer, sheet_name='All_Stores', index = False)
    
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['All_Stores']
    
   

    # Add some cell formats.
    currency_format = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
    #cell_format = workbook.add_format()
    #cell_format.set_align('left')    

    #cell_format = workbook.add_format({'align': 'left'})
    
    #per_format  =  workbook.add_format({'num_format': '0%'})
    
   # worksheet.set_column('A:A',cell_format)#Code
   # worksheet.set_column('B:B',cell_format)#Description
   # worksheet.set_column('C:C',cell_format)#Qty
    #worksheet.set_column('D:D',currency_format)#Sales
   # worksheet.set_column('E:E',cell_format)#Qty
   # worksheet.set_column('F:F',currency_format)#Sales
   # worksheet.set_column('G:G',cell_format)#%Change
   # worksheet.set_column('H:H',cell_format)#Qty
    #worksheet.set_column('I:I',currency_format)#Sales
    #worksheet.set_column(0, last_col - 1, format)

    
    (last_row, last_col) = dfObj.shape
    
    column_settings = [{'header': 'Code', }, 
                       {'header': 'Description', }, 
                       {'header': 'FY_2020_Qty',}, 
                       {'header': 'FY_2020_Sales','format': currency_format,}, #'format':currency_format
                       {'header': 'FY_2019_Qty', }, 
                       {'header': 'FY_2019_Sales','format': currency_format, },
                       {'header': 'Per_Chg_Periods',},
                       {'header': 'YTD_This_Year_Qty'}, 
                       {'header': 'YTD_This_Yr_Sales','format': currency_format,}, 
                       {'header': 'YTD_Last_Year_Qty',}, 
                       {'header': 'YTD_Last_Yr_Sales','format': currency_format,}, 
                       {'header': 'Per_Chg_Yrs', }, 
                       {'header': 'Period', }, 
                       {'header': 'Store_Name', }]

    # Create a list of column headers, to use in add_table().
    #column_settings = [{'header': column} for column in df.columns]
     #Align cells left justified
    #format = workbook.add_format()
    #format.set_align('left')

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, last_row, last_col-1,{'columns': column_settings, 'style':'Blue, Table Style Light 13' }) 
    if values[2] is True:
        worksheet2 = workbook.add_charsheet('Store_Performance')
        #chart = workbook.add_chart({'type': 'column'})
        #chart.add_series({'values': '=All_Stores!$A$1:$A$5'})
        #chart.add_series({'values': '=All_Stores!$B$1:$B$5'})
        #chart.add_series({'values': '=All_Stores!$C$1:$C$5'})

        # Insert the chart into the worksheet.
        #worksheet.insert_chart('A7', chart)

        
   
    

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    # done with loop... need to destroy the window as it's still open
    window.close()
    
        #Popup that tells our users where the files are at
    sg.popup_animated(image_source=None)
    sg.popup('View results at ' +  out_path,keep_on_top=True)

    writer.close()
    
    
   

        




