from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import NamedStyle
from pandas import DataFrame
from openpyxl.styles import Font
import shutil
import pandas as pd
import numpy as np
import os
import time


class Excel:
    __titles = []
    __nrows = 0
    def __init__(self,fromFile,toFile):
        self.fromFile = fromFile
        self.toFile = toFile


    def initialize(self):
        self.wb = load_workbook(self.fromFile)

        main_sheet = 'DENEME MUAVİN 1'

        self.ws = self.wb[main_sheet]
        self.__nrows = len([row for row in self.ws['C'] if row.value != None])


        """ Final sheet """

        self.resultSheet = self.wb.create_sheet("Outcome")
        self.resultSheet.sheet_format.defaultColWidth = 15  # custom width
        self.resultSheet.column_dimensions['B'].width = 40
        self.resultSheet.column_dimensions['C'].width = 15

        """"""

        excelData = list(self.ws.values)  # getting the data inside the sheet
        titles = list(excelData[0]) # getting only the first item (titles)


        self.wb.remove(self.wb[main_sheet]) # deleting the old sheet to be overwritten



        self.df = DataFrame(excelData[1:],columns=titles)
        self.df.rename(columns = {'Şube':"Balance"},inplace=True)
        self.deleteColumns(['Stok Kodu','Sıralama','Belge Türü','Belge Türü Açıklaması',
                            'Masraf Merkezi','COLUMN1','Yevmiye No','Fiş Türü'])

        self.replace_column(['Borç Tutarı','Alacak Tutarı'],np.NAN,0)

        # calcualte the balance => Borc Tutari - Alacak Tutari
        self.df["Balance"] = self.df[['Borç Tutarı','Alacak Tutarı']].apply(lambda x:x[0]-x[1],axis=1)


        self.replace_column(['Borç Tutarı','Alacak Tutarı'],np.NAN,0) # remove nan from the columns


        # apply the function => if balance >=0 return doviz tutari else: return 0
        self.df["PV Döviz"] = self.df[["Balance",'Döviz Tutarı']].apply(self.pv_doviz,axis=1)


        # cleaning columns
        self.replace_column(["PV Döviz"],np.NAN,0)
        self.replace_column(['Hesap Adı'],["None",None],'-')


        # get only the names of the companies
        filtered_names = self.df[['Hesap Kodu','Hesap Adı']][self.df['Hesap Adı']!="-"]

        # get the indexes that includes company names
        delete_indexes = list(filtered_names.index)

        # generate a dictionary Hesap Kodu : Hesap Adi
        names = dict(zip(filtered_names['Hesap Kodu'],filtered_names['Hesap Adı']))

        # fill Hesap Adi with the Hesap Kodu based on the Hesap Kodu
        self.df['Hesap Adı'] = self.df['Hesap Kodu'].apply(lambda x: names[x])

        # delete the rows that had the company names
        self.df.drop(labels=delete_indexes,inplace=True)

        # fill C1 and C3 with the manipulation of Hesap Kodu
        account_codes = list(self.df['Hesap Kodu'])
        self.df.insert(0,'C1',[value[0] for value in account_codes])
        self.df.insert(1,'C3',[value.split(".")[0] for value in account_codes])

        # format the date from long to short date
        self.df[['Fiş Tarihi','Evrak Tarihi']] = self.df[['Fiş Tarihi','Evrak Tarihi']].apply(lambda x: pd.to_datetime(x))



        # recreate the main shet
        self.new_sheet = self.wb.create_sheet(main_sheet,0)
        self.new_sheet.sheet_format.defaultColWidth = 15  # custom width
        self.new_sheet.column_dimensions['D'].width = 40
        self.new_sheet.column_dimensions['I'].width = 40

        # fill the main sheet
        self.saveToSheet(self.df,self.new_sheet)


        results_data_frame = self.df[['Hesap Kodu','Hesap Adı','Balance']]
        results_data_frame = results_data_frame.groupby(['Hesap Kodu','Hesap Adı'],as_index=False)['Balance'].sum()


        # fill the outcome sheet

        self.saveToSheet(results_data_frame,self.resultSheet)
        last_index = str(len(list(results_data_frame['Hesap Kodu']))+2)
        self.resultSheet.merge_cells('A'+last_index+":B"+last_index)
        self.resultSheet['A'+last_index].value = "Grand Total"
        self.resultSheet['C'+last_index].value = results_data_frame['Balance'].sum()
        self.resultSheet['A'+last_index].font = Font(bold=True,size=12)
        self.cells_format(['C'],'Comma',self.resultSheet)



        date_style = NamedStyle(name='datetime', number_format='YYYY-MM-DD')
        self.cells_format(['E','G'],date_style,self.new_sheet)
        self.cells_format(['J','K',"L","N","O",'Q'],'Comma',self.new_sheet)


        self.saveFile()

    def saveToSheet(self,df,toSheet):
        for row in dataframe_to_rows(df,index=False,header=True):
            toSheet.append(row)

        for cell in toSheet['1:1']:
            cell.font = Font(bold=True,size=12)

    def saveFile(self):
        try:
            self.wb.save(self.toFile)
        except:
            print ('in')
            os.system("taskkill /f /im Excel.exe")
            time.sleep(1)
            self.wb.save(self.toFile)

    def pv_doviz(self,x):
        balance = x[0]
        if(balance>=0):
            return x[1]
        else:
            return -x[1]

    def cells_format(self,column_list,formatType,sheet):
        for row in range(2,self.__nrows+1):
            for col in column_list:
                sheet[col+str(row)].style = formatType

    def replace_column(self,column_list,value_remove,value_place):
        for col in column_list:
            self.df[col].replace(value_remove,value_place,inplace=True)


    def deleteColumns(self,names_list):
        for name in names_list:
            del self.df[name]





def calcualte(path):
    direc_files = os.listdir(path)
    save_to = "results"
    save_to_path = os.path.join(path,save_to)
    try:
        if save_to in direc_files:
            shutil.rmtree(save_to_path)

        os.mkdir(save_to_path)

        for file_name in direc_files:
            if(file_name.split(".")[-1]=="xlsx"):

                path_to_save = os.path.join(save_to_path,file_name)
                path_to_file = os.path.join(path,file_name)
                ExcelObject = Excel(path_to_file,path_to_save)
                ExcelObject.initialize()
            return 1
    except:
        return 0

flag  = calcualte(r'C:\Users\Fahed Sabellioglu\Desktop\data')

print (flag)




