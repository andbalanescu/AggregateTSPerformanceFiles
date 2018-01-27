##################



import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

# Performance Files need to have "performance" in filename name; result file will automatically get xlsx extension !

def getDataFrame(fileNameA, osDir):
    os.chdir(osDir)
    dfA1 = pd.read_excel(fileNameA,sheetname="Trades List",skip_footer=1,parse_cols=[0,2,7,9],skiprows=3)
    dfA1.columns = ['#','Datetime','CumNetProfit','Runupdown']
    dfA1 = dfA1[pd.isnull(dfA1['#'])]
    
    dfA1.index = dfA1['Datetime']
    del dfA1['Datetime']
    del dfA1['#']
    #del dfA1['Runupdown']   
    return dfA1;

def processTwoStrategies(dfA1, dfA2 = None):
#    plt.figure()
#    dfA1.plot(subplots=True, title='performance A, native')
    # we get the min PNL in a week (this helps plotting more realistic dips regarding drawdown);
    # as it's a cumulative sum, following week's min it's going to be {max of prev. week, or a new lower low}
    dfA1W = dfA1.resample('W').min()
    dfA1W['CumNetProfit'] = dfA1W['CumNetProfit'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
    dfA1W['Runupdown'] = dfA1W['Runupdown'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
#    plt.figure()
#    dfA1W.plot(subplots=True, title='performance A, weekly aggregation')
    temp = dfA1W;
    
    if dfA2 is not None:
#        plt.figure()
#        dfA2.plot(subplots=True, title='performance B, native')
        dfA2W = dfA2.resample('W').min()
        dfA2W['CumNetProfit'] = dfA2W['CumNetProfit'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
        dfA2W['Runupdown'] = dfA2W['Runupdown'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
#        plt.figure()
#        dfA2W.plot(subplots=True, title='performance B, weekly aggregation')
        
        index1 = dfA1W.index
        index2 = dfA2W.index
        new_index = index1.join(index2, how='outer')
        
        dfA1W = (dfA1W.reindex(index=new_index)).bfill().ffill()
        dfA2W = (dfA2W.reindex(index=new_index)).bfill().ffill()
        
        temp = dfA1W.add(dfA2W);
#        plt.figure()
#        temp.plot(subplots=True, title='Totals of previous two performances')

    return temp;

#temp = processTwoStrategies(getDataFrame(fileNameA, osDir), getDataFrame(fileNameB, osDir))
#temp = processTwoStrategies(getDataFrame(fileNameA, osDir), None)

def iteratePerfromanceFiles(osDir, fileExt = ".xlsx", monthly_file = None):
    tempPerformance = None;
    for file in os.listdir(osDir):
        if file.endswith(".xlsx") and "performance" in file.lower():
            print(os.path.join(osDir, file))
            if tempPerformance is None:
                tempPerformance = processTwoStrategies(getDataFrame(file, osDir), None);
            else:
                tempPerformance = processTwoStrategies(getDataFrame(file, osDir), tempPerformance);
    if tempPerformance is not None:
        plt.figure()
        title = 'All W Combined; MaxDrawdown/MinPNL: ' + str(tempPerformance['Runupdown'].min())
        tempPerformance.plot(subplots=True, title=title, figsize = (14,10))
    else:
        print('NO Tradestation Results FILE FOUND in given Directory!!!')
    if monthly_file is None:
        monthly_file = "aggregated_rezults.xlsx"
    else:
        monthly_file = monthly_file + ".xlsx"
    if monthly_file is not None:
        monthlyRez = tempPerformance.resample('M').min()
        monthlyRez['CumNetProfit'] = monthlyRez['CumNetProfit'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
        monthlyRez['Runupdown'] = monthlyRez['Runupdown'].replace(to_replace=[0,np.nan], method='ffill').to_frame()
        monthlyRez['NetProfit'] = monthlyRez['CumNetProfit'].diff()
        print("Max Drawdown (min PNL): ")
        print(monthlyRez['Runupdown'].min())
        monthlyRez.to_csv(monthly_file)
        plt.figure()
        monthlyRez.plot(subplots=True, title='Monthly Rezults', figsize = (14,10))
        
        picName = "temp.png"
        fig = plt.figure(figsize = (14,10))

        ax = fig.add_subplot(311)
        ax.plot(monthlyRez['CumNetProfit'], label = 'CumNetProfit')
        ax.legend(loc='upper left')
        ax = fig.add_subplot(312)
        ax.plot(monthlyRez['Runupdown'])
        ax.legend(loc='lower left')
        ax = fig.add_subplot(313)
        ax.plot(monthlyRez['NetProfit'])
        ax.legend(loc='upper left')


        fig.savefig(picName)
        writer = pd.ExcelWriter(monthly_file, engine='xlsxwriter')
        monthlyRez.to_excel(writer, sheet_name='Sheet1')
        #workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.insert_image('F10', picName)
        writer.save()
        os.remove(picName)
    return tempPerformance;






# Main Program
    
# if running from a python cli, run:
# chdir("c:\\temp")
# run test.py    
# if running from Spider, open this file and press run file (F5)

    

osDir = "C:\\temp"
fileExt = ".xlsx"
fileOutput = "monthly_aggregate_totals"

fileName = fileOutput + fileExt
if os.path.exists(fileName):
    os.remove(fileName)

rez = iteratePerfromanceFiles(osDir, fileExt, fileOutput)