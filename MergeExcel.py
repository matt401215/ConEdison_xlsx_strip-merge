import os
import pandas as pd
from pathlib import Path

#################################################################################
#
# combine_2_xlsx()
# Combines two xlsx files and saves as saveFile
#
#################################################################################
def combine_2_xlsx(filePath1, filePath2, saveFile):
    # Create data object pf files
    df1 = pd.read_excel(filePath1)
    df2 = pd.read_excel(filePath2)

    # Use this print statement to find the headers for the next two lines
    #print(df2)
    
    # Specify what values you want to combine
    values1 = df1[['account_num','address1', 'address2','date_posted','days','bill_date','elect_kwh_usage','elect_ce_charge','elect_esco','gas_therm_usage','gas_ce_charge','gas_esco','tot_billing','other_charges','balance']]
    values2 = df2[['account_num','address1', 'address2','date_posted','days','bill_date','elect_kwh_usage','elect_ce_charge','elect_esco','gas_therm_usage','gas_ce_charge','gas_esco','tot_billing','other_charges','balance']]

    dataframes = [values1,values2]
    join = pd.concat(dataframes)
    join.to_excel(saveFile)
    


#################################################################################
#
# combine_all():
# Combines all the xlsx files in directory into one file saveFile.
# Starts by combining the first two files into saveFile. Then combines every other
# file with saveFile.
#
#################################################################################
def combine_all(directory, saveFile):
    count = 0
    for i, filename in enumerate(os.listdir(directory)):
        if filename.endswith(".xlsx"):
            count += 1
            print("File #",count)
            if count == 1:
                path1 = os.path.join(directory,filename)
            elif count == 2:
                path2 = os.path.join(directory,filename)
                combine_2_xlsx(path1, path2, saveFile)
            else:
                combine_2_xlsx(os.path.join(directory,filename),saveFile,saveFile)


#################################################################################
#
# main()
#
#################################################################################
def main():

    #directory = '/Users/matthewgilroy/Desktop/xlsx Test Folder/AAA'
    #savePath = '/Users/matthewgilroy/Desktop/xlsx Test Folder/AAA'

    directory = '/Volumes/GoogleDrive/My Drive/Projects/Genesis Realty Group/Bill Auditing/Historical Utility and Supply Data, contracts/Con Edison/Excel Statement Summaries (Final)/Modified Final/Step2'
    savePath = '/Volumes/GoogleDrive/My Drive/Projects/Genesis Realty Group/Bill Auditing/Historical Utility and Supply Data, contracts/Con Edison/Excel Statement Summaries (Final)/Modified Final/Final'


    
    newFile = "ALLorNOTHING4.xlsx"

    saveFile = Path(savePath,newFile)
    
    combine_all(directory, saveFile)

    print("Finnish")

main()



    
