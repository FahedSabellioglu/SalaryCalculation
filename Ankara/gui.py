from tkinter import *
import tkinter.filedialog as tkFileDialog
import tkinter.messagebox as Msg
import os
import re
from operations import Excel

class GUI(Frame):

    __files = []
    def __init__(self,parent):
        Frame.__init__(self,parent)
        self.initGui()

    def initGui(self):

        """App Title"""
        Label(self, text="JKE", bg="red4", fg="white",
              font=("", "12", "bold")).grid(row=0, column=0, columnspan=8, sticky=EW,pady=(0,5))


        "Path Label with an Entry to store user chosen direc"

        Label(self,text="Path:",font=("", "9", "bold")).grid(row=1,column=0,padx=(0,5),pady=(0,5),sticky=W)
        self.DirecPath = Entry(self,state='readonly',width=55)
        self.DirecPath.grid(row=1,column=1,columnspan=6,padx=(0,5),pady=(0,5),ipady=3,sticky=W+E)

        """"""

        """Entry for the names of the final excel workbook"""

        Label(self,text="Save As:",font=("", "9", "bold")).grid(row=2,column=0,sticky=W+E)
        self.saveName = Entry(self)
        self.saveName.insert(END,'Final')
        self.saveName.grid(row=2,column=1,columnspan=3,sticky=W,ipady=3)

        """"""

        notes_frame = LabelFrame(self,text="Pay Attention:",font=("", "10", "bold"))
        notes_frame.grid(row=4,column=0,columnspan=8,sticky=EW,pady=(10,0),padx=(5,5),ipady=4)

        self.labelVar = StringVar()
        self.labelVar.set('Files count: ')
        Label(notes_frame,textvariable=self.labelVar).grid(row=0,column=1,sticky=W)
        self.count = Label(notes_frame,font=("", "10", "bold"))
        self.count.grid(row=0,column=2,sticky=W)

        general_notes="Rules:\n" \
              "1- Make sure you save and close all Excel files before you run the app;\n    Otherwise, You will lose all the unsaved work.\n" \
              " 2- The excel files which will be processed should be in the chosen path.\n" \
              " 3- For further info, Please read the docs." \

        Label(notes_frame,text=general_notes,justify='left').grid(row=1,column=1,columnspan=7,sticky=EW)

        self.__sub_frames(notes_frame,"Balance",2,' a folder name.',0)
        self.__sub_frames(notes_frame,"Pivot",3,' an Excel file name.',1)

        Label(self,text="By IE",font=("", "8", "bold")).grid(row=5,column=0,columnspan=2,sticky=W)
        self.grid()

    def __sub_frames(self,parentFrame,LabelName,row_n,note,cmd_type):
        frame = LabelFrame(parentFrame,text=LabelName+":",font=("", "9", "bold"))
        frame.grid(row = row_n,column = 0, columnspan = 8, sticky=EW, pady=(10,5),padx=(5,5))
        value = 152
        row_value = 1

        button_text=" Choose a Folder"

        if LabelName == 'Balance':
            self.company_name = Entry(frame,width = 25)
            self.balance_date = Entry(frame,width=25)

            self.company_name.insert(END,'JKE Global')
            self.balance_date.insert(END,'06.09.2019')
            Label(frame,text=" Date:").grid(row = 3, column = 0,sticky=W)
            Label(frame,text=" Company Name:").grid(row = 2, column = 0, sticky=W)
            self.company_name.grid(row = 2, column = 1,ipady=3, columnspan = 7,sticky=EW,pady=(0,10),padx=(0,3))
            self.balance_date.grid(row = 3, column = 1,ipady=3, columnspan = 2,pady=(0,5))

            value = 0
            row_value = 4
            button_text = "Choose a File"

        Button(frame,text=button_text,command = lambda : self.ChooseDirec(cmd_type),width = 13).grid(row = row_value , column = 0, columnspan = 2,sticky=W)
        Button(frame,text='Run',command = lambda : self.Run(cmd_type),font=("", "9", "bold"),width=15).grid(row = row_value , column = 7, columnspan = 1, sticky=E+W,padx=(value,0),pady=(5,5))


    def ChooseDirec(self,opType):

        self.count.config(text='')


        if opType == 1:
  
            self.company_name.delete(0,END)
            self.balance_date.delete(0,END)
            self.fileDialog = tkFileDialog.askdirectory(title = 'Select a Directory:')
    
        elif opType == 0:
            self.fileDialog = tkFileDialog.askopenfilename(title = "Select An Excel File",filetypes=[("Excel files", "*.xlsx")])
            self.direc_path = os.path.dirname(self.fileDialog)
            

            self.file_name = os.path.basename(self.fileDialog)
        

        self.DirecPath.config(state='normal')
        self.DirecPath.delete(0,END)
        self.DirecPath.insert(END,self.fileDialog)
        

        if opType == 1:
            self.getFiles()
            self.label_string('Files count: ')
            if self.DirecPath.get().strip() != "":
                if (len(self.__files) == 0):
                    self.count.config(text="No excel files in this directory.")
                else:
                    self.count.config(text=len(self.__files))
        elif opType == 0:
            self.label_string("File name: ")
            self.count.config(text=self.file_name)
        self.DirecPath.config(state='disabled')

    def label_string(self,value):
        self.labelVar.set(value)
        self.update_idletasks()


    def getFiles(self):
        try:
            self.__files = [file for file in os.listdir(self.fileDialog) if file.split('.')[-1]=='xlsx' and "~" not in file and self.check_name(file)==None]
        except FileNotFoundError:
            return
        except AttributeError:
            return

    def Run(self,opType):
        saveName = self.saveName.get()

        if ("." in saveName): # prevent the dots in the save name
            Msg.showerror("Name Error",'Please write a name without " . "')


        elif(self.DirecPath.get()==''):
            if opType == 0:
                Msg.showwarning('Notice',"Please choose a file first.")
            elif opType == 1:
                Msg.showwarning('Notice',"Please choose a folder first.")

        else:
            saveName = self.excel_name()
            saveName += ".xlsx"
            if opType == 1:
                if os.path.isfile(self.fileDialog):
                    Msg.showwarning('Wrong Path',"The chosen path if not compatible with the operation, please check and try again.")
                    return

                elif (len(self.__files)==0):
                    Msg.showwarning("Notice", "The path you chose does not contain any excel files!")

                else:
                    self.update_count_label(len(self.__files))
                    self.calcualte(saveName,1)
            elif opType == 0:
                pattern = re.compile("^[0-9]{2}.[0-9]{2}.[0-9]{4}$")

                if (self.company_name.get().strip()==""):
                    Msg.showwarning("Notice","You can't leave the company name field empty.")

                elif (pattern.search(self.balance_date.get())==None):
                    Msg.showwarning("Notice","Please follow the format DD.MM.YYYY")


                else:
                    self.block_entries('disabled')
                    self.update_count_label(self.file_name+"==>"+saveName)
                    self.calcualte(saveName,0)

    def update_count_label(self,value):
        self.count.config(text=value)

    def check_name(self,to_test):
        patters = re.compile("^"+self.saveName.get()+"\.xlsx$|(^"+self.saveName.get()+".*?\)\.xlsx$)")
        return patters.search(to_test)

    def excel_name(self):

        if os.path.isfile(self.fileDialog):
            NamesCount = len([file for file in os.listdir(os.path.dirname(self.fileDialog)) if (self.check_name(file)!=None)])
        elif os.path.isdir(self.fileDialog):
            NamesCount = len([file for file in os.listdir(self.fileDialog) if (self.check_name(file)!=None)])

        if NamesCount!=0:
            return "{0} ({1})".format(self.saveName.get(),NamesCount)
        elif NamesCount==0:
            return self.saveName.get()

    def block_entries(self,e_state):
        self.balance_date.config(state = e_state)
        self.company_name.config(state = e_state)

    def updateGui(self):

        self.count.config(text="Processing " + str(len(self.__files)) + " verified files.")
        self.update_idletasks()
        self.__files = [os.path.join(self.fileDialog, file) for file in self.__files]

    def save_as(self,savename):
        self.label_string('Status: ')
        self.update_count_label('Saved as ' + savename)
    def calcualte(self,saveName,opType):
        self.saveName.config(state='disabled')
        try:
            if opType == 1:
                self.updateGui()
                ExcelObject = Excel(self.__files, self.fileDialog, saveName)
                ExcelObject.loadFiles()
                

            elif opType == 0:
                self.update_idletasks()
                ExcelObject = Excel(files=[],path= self.fileDialog,saveAs=saveName,b_date=self.balance_date.get(),b_name=self.company_name.get())
                ExcelObject.bilanco()
            self.save_as(saveName)
            Msg.showinfo("Notice","Done, please check your directory.")


        except KeyError as E:
            self.error_state()
            Msg.showwarning('Notice',"The input file "+ str(E) + " does not follow the accepted format.")
        except ValueError as E:
            self.error_state()
            Msg.showwarning('Notice',"The input file "+str(E)+" does not follow the accepted format.")
        except PermissionError as E:
            self.error_state()
            Msg.showwarning('Notice',str(E))

        except Exception as E:            
            self.error_state()
            Msg.showwarning('Error','Error on line {}'.format(sys.exc_info()[-1].tb_lineno)+" "+type(E).__name__+"\n" + str(E) +"\n\nPlease send the input files with a screenshot of this window to get help.")
        finally:
            self.saveName.config(state='normal')
            self.block_entries('normal')

    def error_state(self):
        self.update_count_label('Failed!')
        self.label_string('Status:')


def main():
    root = Tk()
    guiObj = GUI(root)
    root.title("JKE")
    root.resizable(False, False)
    datafile = "rocket+1.ico"
    if not hasattr(sys, "frozen"):
        datafile = os.path.join(os.path.dirname(__file__), datafile)
    else:
        datafile = os.path.join(sys.prefix, datafile)
    root.iconbitmap(default=datafile)
    root.mainloop()


if __name__=='__main__':
    main()
