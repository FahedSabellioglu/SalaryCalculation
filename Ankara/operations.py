from openpyxl import load_workbook,Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import DataFrame
from openpyxl.styles import Font
from numpy import NaN as NaN
import os
import time
import excel
import datetime

__author__ = "Fahed Sabellioglu"

class Excel:
    __nrows = 0

    def __init__(self,files,path,saveAs):
        self.files = files
        self.path = path
        self.saveAs = saveAs
        self.dfs = DataFrame()


    def loadFiles(self):
        counter = 1
        for file in self.files:
            open_workbook = load_workbook(filename=file,read_only=True)
            sheet_name =open_workbook.sheetnames[0]

            ws = open_workbook[sheet_name]

            excelData = list(ws.values)  # getting the data inside the sheet
            titles = list(excelData[0]) # getting only the first item (titles)

            # creating temporary dataframe
            df = DataFrame(excelData[1:],columns=titles)

            # renaming sube to balance
            df.rename(columns = {'Şube':"Balance"},inplace=True)

            # dropping the unneeded columns
            try:
                self.deleteColumns(['Stok Kodu','Sıralama','Belge Türü','Belge Türü Açıklaması',
                                    'Masraf Merkezi','COLUMN1','Yevmiye No','Fiş Türü'],df)
            except KeyError:
                open_workbook.close()
                raise KeyError


            df.insert(0,'identifier',counter)

            #replace 'None' or None with -
            self.replace_column(['Hesap Adı'],["None",None],'-',df)

            # get only the names of the companies
            filtered_names = df[['Hesap Kodu','Hesap Adı']][df['Hesap Adı']!="-"]

            # get the indexes that includes company names
            delete_indexes = list(filtered_names.index)


            # generate a dictionary Hesap Kodu : Hesap Adi
            names = dict(zip(filtered_names['Hesap Kodu'],filtered_names['Hesap Adı']))

            # fill Hesap Adi with the Hesap Kodu based on the Hesap Kodu
            df['Hesap Adı'] = df['Hesap Kodu'].apply(lambda x: names[x])

            # delete the rows that had the company names
            df.drop(labels=delete_indexes,inplace=True)


            # append the temporary dataframe to the main dataframe
            self.dfs = self.dfs.append(df,ignore_index=True)

            counter += 1

        self.__nrows = len(self.dfs.index)

        # change NaNs into 0 to do the calculations
        self.replace_column(['Borç Tutarı', 'Alacak Tutarı'], NaN, 0,self.dfs)

        # performing Balance = Borc Tutari - Alacak Tutari
        self.dfs["Balance"] = self.dfs[['Borç Tutarı','Alacak Tutarı']].apply(lambda x:x[0]-x[1],axis=1)

        self.replace_column(['Borç Tutarı','Alacak Tutarı'],NaN,0,self.dfs) # remove nan from the columns

        # apply the function => if balance >=0 return doviz tutari else: return 0
        self.dfs["PV Döviz"] = self.dfs[["Balance",'Döviz Tutarı']].apply(self.pv_doviz,axis=1)

        #replacing Nans with zero
        self.replace_column(["PV Döviz"],NaN,0,self.dfs)

        # fill C1 and C3 with the manipulation of Hesap Kodu
        account_codes = list(self.dfs['Hesap Kodu'])
        self.dfs.insert(1,'C1',[value[0] for value in account_codes])
        self.dfs.insert(2,'C3',[value.split(".")[0] for value in account_codes])

        # format the date from long to short date
        self.dfs[['Fiş Tarihi','Evrak Tarihi']] = self.dfs[['Fiş Tarihi','Evrak Tarihi']].apply(lambda x: x.dt.strftime('%d/%m/%Y'))

        self.dfs = self.dfs.reset_index(drop=True)

        final_workbook = Workbook()
        ws = final_workbook['Sheet']
        ws.title = 'Outcome'

        ws.sheet_format.defaultColWidth = 15  # custom width
        ws.column_dimensions['E'].width = 40
        ws.column_dimensions['J'].width = 30

        # save the sheet
        self.saveToSheet(self.dfs,ws)

        self.cells_format(['K',"L",'M',"O",'P'],'Comma',ws)

        path = os.path.join(self.path,"Results.xlsx")

        self.saveFile(final_workbook,path)

        self.CreatePivot(path)

    def CreatePivot(self,file_path):

        excel.Pivot(self.path,self.saveAs)

        os.remove(file_path)

    def saveToSheet(self,df,toSheet):

        for row in dataframe_to_rows(df,index=False,header=True):
            toSheet.append(row)

        for cell in toSheet['1:1']:
            cell.font = Font(bold=True,size=12)

    def saveFile(self,ws,path):
        try:
            ws.save(path)
        except:
            os.system("taskkill /f /im Excel.exe")
            time.sleep(0.5)
            ws.save(path)

    def pv_doviz(self,x):
        balance = x[0]
        if(balance>=0):
            return x[1]
        else:
            return -x[1]

    def time_formart(self,sheet):
        for row in range(2,self.__nrows+1):
            for col in ['F','H']:
                dttm = datetime.datetime.strptime(sheet[col+str(row)].value, "%d/%m/%Y")
                sheet[col+str(row)].value = dttm

    def cells_format(self,column_list,formatType,sheet):
        for row in range(2,self.__nrows+1):
            for col in column_list:
                sheet[col+str(row)].style = formatType

    def replace_column(self,column_list,value_remove,value_place,df):
        for col in column_list:
            df[col].replace(value_remove,value_place,inplace=True)


    def deleteColumns(self,names_list,df):
        for name in names_list:
            del df[name]






