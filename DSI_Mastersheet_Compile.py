# -*- coding: utf-8 -*-
"""
Created on Mon May  9 09:35:03 2022

@author: Duvallp
"""

import pandas as pd
import xlwings as xw

dsi_sheet = "T:/amri/DTIanalysis/DSI/b-table-testing/DSI_File_Find/DSI_List.xlsx"
dsi=pd.ExcelFile(dsi_sheet).parse('DSI_subs')
#df.info

#need to use xlwings to open password protected excel files
amri_sheet = "T:/amri/AMRI_SM_Master_Sheet.xlsx"
amri_sheet=xw.Book(amri_sheet).sheets['Sheet1']
amri = amri_sheet['A1:BO2271'].options(pd.DataFrame, index=False, header=True).value

#remove all bad filetype names
amri=amri.loc[amri['AMRI_SM_Master_File_ID'].str.len()==22]
amri['mri_dos']=amri['AMRI_SM_Master_File_ID'].apply(lambda x: x[:15])

df_keep = amri[amri.mri_dos.isin(dsi.Date_of_Scan)]

df_dsi_ms = df_keep[['MRI', 'SM', 'Sex', 'DX1', 'DOS', 'DOB', 'Age', 'mri_dos', \
                     'Scanner',  'Scanner Model Script', 'Software Version Script', \
                         'Receiving Coil', 'MPRAGE', 'Checked MPRAGE Quality', \
                             'DE60 or 70', 'DTI', 'DSI', 'REV FAT']]

#print("Printing DSI Spreadsheet")
#excel_file = "T:/amri/DTIanalysis/NIFTI/New_DSI_Mastersheet.xlsx"
#df_dsi_ms.to_excel(excel_file, sheet_name='DSI_subs', index=False)


#find what subjects are in the DSI list but not in the AMRI mastersheet
#df_missing = dsi[~dsi.DSI_Sub_DOS.isin(df_keep.mri_dos)]

                 