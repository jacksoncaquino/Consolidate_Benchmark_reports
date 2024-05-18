# -*- coding: utf-8 -*-
"""
Created on Wed Sep  7 21:42:48 2022

@author: Jackson_Aquino
"""

"""
Run a Review compensation benchmark data report exporting to a CSV file
Include survey results, target and spread
Include # incumbents direct and rollup
Include 50th, 90th, 75th percentile
If you have any Excel files, convert them to CSV first

Replace the directories below for the ones you'll use in your computer instead
"""
import pandas
import numpy
from datetime import date
import os
import time

FilesInCurrentFolder = os.listdir("C:\\Users\\windows_user_name\\test\\")
countReadableFiles= 0
csvFiles = []
excelFiles = []
for eachFile in FilesInCurrentFolder:
    if eachFile.split(".")[len(eachFile.split("."))-1].lower() == "csv":
        csvFiles.append(eachFile)
    #if eachFile.split(".")[len(eachFile.split("."))-1].lower() == "xlsx":
    #    excelFiles.append(eachFile)

countReadableFiles = len(csvFiles) #+ len(excelFiles)
if countReadableFiles > 0:
    print("Found " + str(countReadableFiles) + " file(s) to read:")
    for eachFile in csvFiles:
        print(eachFile)

for eachFile in FilesInCurrentFolder:
    if eachFile.split(".")[len(eachFile.split("."))-1].lower() == "xlsx":
        excelFiles.append(eachFile)
    
for eachFile in excelFiles:
    if "Job Profile with Compensation Ranges" in eachFile:
        print("Found the WD Report for job profile lookup: ", eachFile)
        jobProfilesLookupFile = eachFile
        
#print("Please type 0 or 1 to select the population for which you would like to create the benchmark file:")
execSelection = "0"
#while execSelection !="0" and execSelection != "1":
#    execSelection = input("Non-Execs: 0, Execs: 1 ==> ")

#get current time
timestart = time.time()
print(timestart)
#create the lookup table for job profile information:

jobProfilesLookup = pandas.read_excel(jobProfilesLookupFile,sheet_name="Sheet1",skiprows=1)
jobProfilesLookup = jobProfilesLookup.filter(items=['Job Code', 'Job Profile','Management Level', 'Job Family Group', 'Job Family', 'Sales Indicator'])
jobProfilesLookup = jobProfilesLookup.drop_duplicates(subset=["Job Profile"])

countFilesProcessed = 0
for eachFile in csvFiles:
    arquivo = open("C:\\Users\\windows_user_name\\test\\" + eachFile)
    print(eachFile)
    csvtable = pandas.read_csv(arquivo,skiprows=18)
    newtable = csvtable.filter(items=['Job Profile', 'Survey Job', 'Benchmark Profile',
           'Currency', '# Incumbents - Direct', '# Incumbents - Rollup', '50th', '90th', '75th', 'Target Percentile',
           '# Incumbents - Direct.1', '# Incumbents - Rollup.1', '50th.1',
           '90th.1', '75th.1','Midpoint','Midpoint1'])
    newtableNoInactives = newtable[newtable["Job Profile"].str.contains("(inactive)")==False]
    newtableNoInactivesNoBlankCurrency = newtableNoInactives.dropna(subset=["Currency"])
    newtableNoInactivesNoBlankCurrency = newtableNoInactivesNoBlankCurrency.replace(") - Non-Exec (All Job Profiles)",regex=True)
    newtableNoInactivesNoBlankCurrency["Comp Market"] = newtableNoInactivesNoBlankCurrency["Benchmark Profile"].str.split(")",expand=True)[0]
    newtableNoInactivesNoBlankCurrency["Comp Market"] = newtableNoInactivesNoBlankCurrency["Comp Market"].str.split("(",expand=True)[1]
    newtableNoInactivesNoBlankCurrency["Country"] = newtableNoInactivesNoBlankCurrency["Benchmark Profile"].str.split(" \(",expand=True)[0]
    newtableNoInactivesNoBlankCurrency["SearchKey"] = newtableNoInactivesNoBlankCurrency["Job Profile"] + newtableNoInactivesNoBlankCurrency["Comp Market"]
    
    
    """
    #remove duplicates from WTW South Africa
    newtableNoInactivesNoBlankCurrencyp1 = newtableNoInactivesNoBlankCurrency[newtableNoInactivesNoBlankCurrency["Country"] != "South Africa"]
    newtableNoInactivesNoBlankCurrencyp2 = newtableNoInactivesNoBlankCurrency[newtableNoInactivesNoBlankCurrency["Country"] == "South Africa"]
    newtableNoInactivesNoBlankCurrencyp2['dkey'] = newtableNoInactivesNoBlankCurrencyp2['Job Profile'] + newtableNoInactivesNoBlankCurrencyp2['Survey Job']
    newtableNoInactivesNoBlankCurrencyp2 = newtableNoInactivesNoBlankCurrencyp2.drop_duplicates(subset=['dkey'])
    newtableNoInactivesNoBlankCurrencyp2 = newtableNoInactivesNoBlankCurrencyp2.drop(columns=['dkey'])
    newtableNoInactivesNoBlankCurrency = pandas.concat([newtableNoInactivesNoBlankCurrencyp1,newtableNoInactivesNoBlankCurrencyp2])
    """
    
    newtableNoInactivesNoBlankCurrency = newtableNoInactivesNoBlankCurrency.drop(columns=['Survey Job'])
            
    #create another df only with exclusive job profiles, showing the percentiles of each
    lookupTable = newtableNoInactives.filter(items=['Job Profile','Target Percentile'])
    lookupTable = lookupTable.dropna(subset=["Target Percentile"])
    lookupTable = lookupTable.drop_duplicates(subset=["Job Profile"])
    
    #Use this lookup table to feed into a new column on the main table, as a vlookup
    newtableNoInactivesNoBlankCurrency = pandas.merge(newtableNoInactivesNoBlankCurrency, lookupTable,on="Job Profile")
    
    #Use numpy where to determine which column to use for the midpoint on each row: https://www.dataquest.io/blog/tutorial-add-column-pandas-dataframe-based-on-if-else-condition/        newtableNoInactivesNoBlankCurrency["Base - Midpoint"] = numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="75th",newtableNoInactivesNoBlankCurrency["75th"],numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="90th",newtableNoInactivesNoBlankCurrency["90th"],newtableNoInactivesNoBlankCurrency["50th"]))
    newtableNoInactivesNoBlankCurrency["Base - Midpoint"] = numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="75th",newtableNoInactivesNoBlankCurrency["75th"],numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="90th",newtableNoInactivesNoBlankCurrency["90th"],newtableNoInactivesNoBlankCurrency["50th"]))
    newtableNoInactivesNoBlankCurrency["TTC - Midpoint"] = numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="75th",newtableNoInactivesNoBlankCurrency["75th.1"],numpy.where(newtableNoInactivesNoBlankCurrency["Target Percentile_y"]=="90th",newtableNoInactivesNoBlankCurrency["90th.1"],newtableNoInactivesNoBlankCurrency["50th.1"]))
    
    #pra testar:
    #newtableNoInactivesNoBlankCurrency.to_csv("Teste.csv",index=False)

    #Use numpy to remove the number of incumbents where the column with the numbers returns a zero
    newtableNoInactivesNoBlankCurrency['# Incumbents - Direct'] = numpy.where(newtableNoInactivesNoBlankCurrency['Base - Midpoint'].isnull(),0,newtableNoInactivesNoBlankCurrency['# Incumbents - Direct'])
    newtableNoInactivesNoBlankCurrency['# Incumbents - Rollup'] = numpy.where(newtableNoInactivesNoBlankCurrency['Base - Midpoint'].isnull(),0,newtableNoInactivesNoBlankCurrency['# Incumbents - Rollup'])
    newtableNoInactivesNoBlankCurrency['# Incumbents - Direct.1'] = numpy.where(newtableNoInactivesNoBlankCurrency['TTC - Midpoint'].isnull(),0,newtableNoInactivesNoBlankCurrency['# Incumbents - Direct.1'])
    newtableNoInactivesNoBlankCurrency['# Incumbents - Rollup.1'] = numpy.where(newtableNoInactivesNoBlankCurrency['TTC - Midpoint'].isnull(),0,newtableNoInactivesNoBlankCurrency['# Incumbents - Rollup.1'])


    #table to sum # of incumbents
    incsTable = newtableNoInactivesNoBlankCurrency.filter(items=["SearchKey",'# Incumbents - Direct','# Incumbents - Rollup','# Incumbents - Direct.1', '# Incumbents - Rollup.1'])
    incsTable = incsTable.groupby("SearchKey").sum()
    
    #incorporate the new counts back on the table
    newtableNoInactivesNoBlankCurrency = pandas.merge(newtableNoInactivesNoBlankCurrency,incsTable,on="SearchKey",suffixes=("","_Total"))
    
    #calculate weights
    newtableNoInactivesNoBlankCurrency["Base direct - weight"] = newtableNoInactivesNoBlankCurrency["# Incumbents - Direct"]/newtableNoInactivesNoBlankCurrency["# Incumbents - Direct_Total"]
    newtableNoInactivesNoBlankCurrency["Base direct - weight"] = newtableNoInactivesNoBlankCurrency["Base direct - weight"].round(3)
    
    newtableNoInactivesNoBlankCurrency["TTC direct - weight"] = newtableNoInactivesNoBlankCurrency["# Incumbents - Direct.1"]/newtableNoInactivesNoBlankCurrency["# Incumbents - Direct.1_Total"]
    newtableNoInactivesNoBlankCurrency["TTC direct - weight"] = newtableNoInactivesNoBlankCurrency["TTC direct - weight"].round(3)
    
    newtableNoInactivesNoBlankCurrency["Base Rollup - weight"] = newtableNoInactivesNoBlankCurrency["# Incumbents - Rollup"]/newtableNoInactivesNoBlankCurrency["# Incumbents - Rollup_Total"]
    newtableNoInactivesNoBlankCurrency["Base Rollup - weight"] = newtableNoInactivesNoBlankCurrency["Base Rollup - weight"].round(3)
    
    newtableNoInactivesNoBlankCurrency["TTC Rollup - weight"] = newtableNoInactivesNoBlankCurrency["# Incumbents - Rollup.1"]/newtableNoInactivesNoBlankCurrency["# Incumbents - Rollup.1_Total"]
    newtableNoInactivesNoBlankCurrency["TTC Rollup - weight"] = newtableNoInactivesNoBlankCurrency["TTC Rollup - weight"].round(3)
    
    #Multiply #weights by the midpoints
    newtableNoInactivesNoBlankCurrency["Base - direct Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["Base - Midpoint"] * newtableNoInactivesNoBlankCurrency["Base direct - weight"]
    newtableNoInactivesNoBlankCurrency["TTC - direct Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["TTC - Midpoint"] * newtableNoInactivesNoBlankCurrency["TTC direct - weight"]
    newtableNoInactivesNoBlankCurrency["Base - rollup Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["Base - Midpoint"] * newtableNoInactivesNoBlankCurrency["Base Rollup - weight"]
    newtableNoInactivesNoBlankCurrency["TTC - rollup Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["TTC - Midpoint"] * newtableNoInactivesNoBlankCurrency["TTC Rollup - weight"]
    
    newtableNoInactivesNoBlankCurrency["Base - direct Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["Base - direct Midpoint weighted"].round(2)
    newtableNoInactivesNoBlankCurrency["TTC - direct Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["TTC - direct Midpoint weighted"].round(2)
    newtableNoInactivesNoBlankCurrency["Base - rollup Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["Base - rollup Midpoint weighted"].round(2)
    newtableNoInactivesNoBlankCurrency["TTC - rollup Midpoint weighted"] = newtableNoInactivesNoBlankCurrency["TTC - rollup Midpoint weighted"].round(2)
    
    #group by to sum by searchkey
    sumTable = newtableNoInactivesNoBlankCurrency.filter(items=['SearchKey', 'Base - direct Midpoint weighted', 'TTC - direct Midpoint weighted', 'Base - rollup Midpoint weighted','TTC - rollup Midpoint weighted'])
    sumTable = sumTable.groupby("SearchKey").sum()
    sumTable['Base - direct Midpoint weighted'] = numpy.where(sumTable['Base - direct Midpoint weighted']==0,"",sumTable['Base - direct Midpoint weighted'])
    sumTable['TTC - direct Midpoint weighted'] = numpy.where(sumTable['TTC - direct Midpoint weighted']==0,"",sumTable['TTC - direct Midpoint weighted'])
    sumTable['Base - rollup Midpoint weighted'] = numpy.where(sumTable['Base - rollup Midpoint weighted']==0,"",sumTable['Base - rollup Midpoint weighted'])
    sumTable['TTC - rollup Midpoint weighted'] = numpy.where(sumTable['TTC - rollup Midpoint weighted']==0,"",sumTable['TTC - rollup Midpoint weighted'])
    
    #put together the final table
    finalTable = newtableNoInactivesNoBlankCurrency.drop_duplicates("SearchKey")
    finalTable = finalTable.filter(items=['Job Profile', 'SearchKey', 'Currency','Comp Market', 'Country', 'Target Percentile_y', '# Incumbents - Direct_Total', '# Incumbents - Rollup_Total', '# Incumbents - Direct.1_Total', '# Incumbents - Rollup.1_Total','Midpoint'])
    finalTable = pandas.merge(finalTable,sumTable,on="SearchKey")
    
    #use another merge to pull the final table
    finalTable = pandas.merge(finalTable, jobProfilesLookup, on="Job Profile")
    
    #rename and reorder columns to mirror the final file
    finalTable = finalTable.rename(columns={'Target Percentile_y':'Target Percentile', 'Management Level': 'Level','# Incumbents - Direct_Total':'Base - # Incs - Direct', '# Incumbents - Rollup_Total':'Base - # Incs - Rollup', '# Incumbents - Direct.1_Total':'TTC - # Incs - Direct', '# Incumbents - Rollup.1_Total':'TTC - # Incs - Rollup', 'Base - direct Midpoint weighted':'Base - Midpoint - Direct', 'TTC - direct Midpoint weighted':'TTC - Midpoint - Direct', 'Base - rollup Midpoint weighted':'Base - Midpoint - Rollup', 'TTC - rollup Midpoint weighted':'TTC - Midpoint - Rollup'})
    finalTable = finalTable[['Job Code', 'Job Profile', 'Country', 'Comp Market', 'Level', 'Job Family Group', 'Job Family', 'Sales Indicator', 'Currency', 'Target Percentile', 'Base - # Incs - Direct', 'Base - Midpoint - Direct', 'TTC - # Incs - Direct', 'TTC - Midpoint - Direct', 'Base - # Incs - Rollup', 'Base - Midpoint - Rollup', 'TTC - # Incs - Rollup', 'TTC - Midpoint - Rollup']]
    
    execLevels = ["E1","E2","E3","E4","E5","I12","I13"]
    
    #filtering for execs only
    if execSelection == "1":
        finalTable = finalTable[finalTable["Level"].isin(execLevels)==True]
    else:
        finalTable = finalTable[finalTable["Level"].isin(execLevels)==False]
    
    if countFilesProcessed == 0:
        finalFile = finalTable
    else:
        finalFile = finalFile.append(finalTable)
    
    countFilesProcessed = countFilesProcessed + 1
    
#Save final file
today = date.today()
print("countFilesProcessed: " + str(countFilesProcessed))
if execSelection == "1":
    execWord = "Execs"
else:
    execWord = "Non-Execs"
    
#finalFileName = "C:\\Users\\windows_user_name\\Downloads\\Global Benchmark File_" + execWord + "_" + today.strftime("%b-%d-%Y") + ".csv"
finalFileName = "C:\\Users\\windows_user_name\\test\\Output\\Global Benchmark File_" + execWord + "_" + today.strftime("%b-%d-%Y") + ".csv"
finalFile.to_csv(finalFileName,index=False)

print("Process finished in " + str(time.time() - timestart) + " seconds")
print("file saved to " + finalFileName)
# input()