from tkinter import *
import tkinter.filedialog as tkFileDialog
import tkinter.messagebox as Msg
import os
import win32com.client
from pywintypes import com_error
import time
from operations import Excel

class GUI(Frame):

    __files = []
    def __init__(self,parent):
        Frame.__init__(self,parent)
        self.initGui()

    def initGui(self):

        """App Title"""
        Label(self, text="JKE", bg="springgreen2", fg="white",
              font=("", "12", "bold")).grid(row=0, column=0, columnspan=8, sticky=EW,pady=(0,5))


        "Path Label with an Entry to store user chosen direc"

        Label(self,text="Path:").grid(row=1,column=1,padx=(0,5),pady=(0,5),sticky=W)
        self.DirecPath = Entry(self,state='readonly',width=40)
        self.DirecPath.grid(row=1,column=2,columnspan=4,padx=(0,5),pady=(0,5),ipady=3)
        Button(self,text="Choose A Path",command = self.ChooseDirec).grid(row=1,column=7,pady=(0,5),padx=(0,5))

        """"""

        """Entry for the names of the final excel workbook"""

        Label(self,text="Save As:").grid(row=2,column=1)
        self.saveName = Entry(self,width=20)
        self.saveName.insert(END,'Final')
        self.saveName.grid(row=2,column=2,columnspan=3,sticky=W,ipady=3)

        """"""
        Button(self,text="Run",command = self.Run).grid(row=2,column=7,columnspan=2,sticky=EW,padx=(0,5))
        notes_frame = LabelFrame(self,text="Pay Attention:")
        notes_frame.grid(row=3,column=0,columnspan=8,sticky=EW,pady=(5,5),padx=(5,5))

        Label(notes_frame,text="Files Count: ").grid(row=0,column=1,sticky=W)
        self.count = Label(notes_frame)
        self.count.grid(row=0,column=2,sticky=W)


        notes="Please pay attention to the followings:\n" \
              "1- Make sure you save all Excel files before you run the app; Otherwise, \n   You will lose the unsaved work.\n"\
              "2- The excel files which will be processed should be in the chosen path.\n" \
              "3- For further info, Please read the docs"


        Label(notes_frame,text=notes,justify='left').grid(row=1,column=1,columnspan=7,sticky=EW)


        Label(self,text="By IE",font=("", "8", "bold")).grid(row=4,column=0,columnspan=2,sticky=W)


        self.grid()

    def ChooseDirec(self):
        self.fileDialog = tkFileDialog.askdirectory(title="Select a Directory")
        self.DirecPath.config(state='normal')
        self.DirecPath.delete(0,END)
        self.DirecPath.insert(END,self.fileDialog)
        self.getFiles()
        self.DirecPath.config(state='disabled')
        if (len(self.__files)==0):
            self.count.config(text="No excel files in this directory.")
        else:
            self.count.config(text=len(self.__files))

    def getFiles(self):
        try:
            self.__files = [file for file in os.listdir(self.fileDialog) if file.split('.')[-1]=='xlsx' and "~" not in file]
        except FileNotFoundError:
            return

    def Run(self):
        saveName = self.saveName.get()
        self.getFiles()
        if ('.' in saveName):
            Msg.showerror("Name Error",'Please Write A Name Without .')

        if saveName in self.__files:
            self.delete_excel(saveName)

        if(self.DirecPath.get()==''):
            Msg.showwarning('Notice',"Please Choose A Path First.")

        elif (len(self.__files)==0):
            Msg.showwarning("Notice","The Path you Chose Does Not Contain Any Excel Files.")

        else:
            self.calcualte(saveName)

    def excelCheck(self):
        try:
            win32com.client.GetActiveObject("Excel.Application")
            os.system("taskkill /f /im Excel.exe")
        except com_error:
            pass

    def delete_excel(self, name):
        file_path = os.path.join(self.fileDialog, name)
        os.remove(file_path)
        self.__files.remove(name)

    def calcualte(self,saveName):
        self.excelCheck()
        time.sleep(0.5)
        name = saveName+".xlsx"
        try:
            if 'Results.xlsx' in self.__files:
                self.delete_excel('Results.xlsx')
            if name in self.__files:
                self.delete_excel(name)
            self.count.config(text="Processing "+str(len(self.__files))+" verified files.")
            self.update_idletasks()

            self.__files = [os.path.join(self.fileDialog,file) for file in self.__files]
            ExcelObject = Excel(self.__files, self.fileDialog,self.saveName.get())
            ExcelObject.loadFiles()
            Msg.showinfo("Notice","Done, Please check your directory.")
        except KeyError:
            Msg.showwarning('Notice',"The path includes files not following the accepted format.")
            self.excelCheck()
        except PermissionError:
            Msg.showinfo("Please do not open the excel files in the chosen directory while the program is working. \nRerun the program")
        except Exception as E:
            Msg.showwarning('Error',E)



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
