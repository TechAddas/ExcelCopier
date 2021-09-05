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
        if inputobject is not None:
            return inputobject.get() is not None and inputobject.get() != ""
        else:
            return ""
    
    def uploadFiles(self):
        if self.hasValue(self.input1_file) and self.hasValue(self.filter_file):
            mc.createDfFromExcel(self.input1_file,self.filter_file,"Site ID" ,1,"R-Inputs")

        if (self.hasValue(self.input2_file) and self.hasValue(self.filter_file)):
            mc.createDfFromExcel(self.input2_file,self.filter_file,"Site ID" ,2,"IP_Input_1")
            mc.createDfFromExcel(self.input2_file,self.filter_file,"Site ID" ,3,"IP_Input_2")

