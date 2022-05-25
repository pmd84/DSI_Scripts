# -*- coding: utf-8 -*-
"""
Created on Wed May  4 13:49:34 2022

@author: Duvallp
"""
import os
import pandas as pd
import numpy as np
from pandas import DataFrame
import xlwt
from xlwt import Workbook
import xlwings as xw

#Import AMRI data

#need to use xlwings to open password protected excel files
amri_sheet = "T:/amri/AMRI_SM_Master_Sheet.xlsx"
amri_sheet=xw.Book(amri_sheet).sheets['Sheet1']
amri = amri_sheet['A1:BO2271'].options(pd.DataFrame, index=False, header=True).value

#remove all bad filetype names                        
amri=amri.loc[amri['AMRI_SM_Master_File_ID'].str.len()==22]
amri['mri_dos']=amri['AMRI_SM_Master_File_ID'].apply(lambda x: x[:15])

data_dir = "T:/amri/DATA/Children/3T"
os.chdir(data_dir)

DSI_dir=[]
DSI_subs=[]
dos_list=[]
DTI=[]
DET2=[]

subjects = os.listdir(data_dir)

for subs in subjects:
    if len(subs) == 4:
        print(subs)
        #os.chdir(data_dir + '/' + subs)
        DOS = os.listdir(data_dir + '/' + subs)
        
        for date in DOS:
            #os.chdir(data_dir + '/' + subs + '/' + date)
            if len(date) == 10:
                folders = os.listdir(data_dir + '/' + subs + '/' +  date) 
            
                if any('DSI' in fold for fold in folders):
                    print("DSI files found for subject", subs, "on date", date)
                    DSI_dir.append(data_dir + '/' + subs + '/' +  date)
                    DSI_subs.append(subs)
                    dos_list.append(date)
                    if any('DTI' in fold1 for fold1 in folders):
                        DTI.append('Y')
                        print("DTI files found")
                    else:
                        DTI.append('N')
                    if any('DET2' in fold2 for fold2 in folders):
                        DET2.append('Y')
                        print("DET2 files found")
                    else:
                        DET2.append('N')
                        
#create dataframe from dsi subjects
dsi = DataFrame({'MRI': DSI_subs, 'Date_of_Scan': dos_list, 'File Path': DSI_dir, 'DTI': DTI, 'DET2': DET2})

#Create an mri_DOS column within the DSI dataframe
dsi['dos'] = dsi['Date_of_Scan']
dsi['dos'].replace('/','_')
dsi['mri_dos'] = dsi.MRI.str.cat(dsi.dos, sep='_')

#compare full path (MRI#_DOS) and only keep amri data for 
amri_keep = amri[amri.mri_dos.isin(dsi.mri_dos)]

#Use these commands to rename or drop columns if you mess up
#dsi = dsi.drop(['DTI (from Mastersheet)','DET2 (from Mastersheet)','REV FAT (from Mastersheet'], 1)
#dsi.rename(columns = {'Subject':'MRI'}, inplace = True)

#Use this command to pull whatever data from the AMRI mastersheet you want
dsi = pd.merge(dsi, amri_keep[['MRI','DTI','DE60 or 70','REV FAT']], on='MRI', how='outer')

excel_file = "T:/amri/DTIanalysis/DSI/b-table-testing/DSI_File_Find/DSI_List.xlsx"
dsi.to_excel(excel_file, sheet_name='DSI_subs', index=False)

