from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename 
import time
import pandas as pd
import MyCopier as mc
import mywritter as mw

class Upload():

    def __init__(self) -> None:
        self.input1_file =  StringVar()
        self.input2_file =  StringVar()
        self.filter_file =  StringVar()
        # mw.removeSheet("Output_Sheet_original.xlsx","output_1_input_A")
        # mw.removeSheet("Output_Sheet_original.xlsx","output_2_input_A")
        # mw.removeSheet("Output_Sheet_original.xlsx","output_2_input_B")
    
    def open_file(self, file):
        file_path = askopenfilename(initialdir = "/", title = "Select a File", filetypes=[('Excel Files', '*xlsx'),("all files")])
        if file_path is not None:
            if file == 'i':
                self.input1_file.set(file_path)
            elif file == 'o':
                self.input2_file.set(file_path)
            elif file == 'f':
                self.filter_file.set(file_path)
            else:
                raise("Something went wrong")
        else:
            print("Please upload a valid file")
        
        print("self.input1_file,self.input2_file,self.filter_file",self.input1_file.get(),self.input2_file.get(),self.filter_file.get())
    
    def hasValue(self, inputobject):
        return inputobject.get() is not None and inputobject.get() != ""
    
    def uploadFiles(self):
        mc.createDfFromExcel(self.input1_file,self.filter_file,"OutPutColumn_Input_1_sheet_1","Site ID",True,"output_1_input_A" ,"R-Inputs")
        mc.createDfFromExcel(self.input2_file,self.filter_file,"OutPutColumn_Input_2_sheet_1","Site ID",True,"output_2_input_A", "IP_Input_1")
        mc.createDfFromExcel(self.input2_file,self.filter_file,"OutPutColumn_Input_2_sheet_2","Site ID",True,"output_2_input_B" ,"IP_Input_2")

