import tkinter as tk
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
import time
from numpy import DataSource 
from upload import Upload as up

class Application():

    def __init__(self,ws):
        self.up = up()
        self.ws=ws
        ws.title('Excel Worker')
        ws.geometry('400x200') 
        self.create_widgets(ws)

    def create_widgets(self,ws):
        input = Label(ws, text='Upload first Input file in xlsx format')
        input.grid(row=0, column=0, padx=10)

        inputbtn = Button(ws, text ='Choose File', command = lambda:self.up.open_file('i',self.show_label_for_file_upload)) 
        inputbtn.grid(row=0, column=1)

        output = Label(ws, text='Upload secomd Input file in xlsx format')
        output.grid(row=1, column=0, padx=10)

        outputbtn = Button(ws, text ='Choose File ', command = lambda:self.up.open_file('o',self.show_label_for_file_upload)) 
        outputbtn.grid(row=1, column=1)

        filter = Label(ws,text='Upload filter data excel in xlsx format ')
        filter.grid(row=2, column=0, padx=10)

        filterbtn = Button(ws, text ='Choose File', command = lambda:self.up.open_file('f',self.show_label_for_file_upload)) 
        filterbtn.grid(row=2, column=1)

        upld = Button(ws, text='Upload Files', command=self.uploadFiles)
        upld.grid(row=3, columnspan=3, pady=10)

    def uploadFiles(self):
        if (self.up.hasValue(self.up.input1_file) and self.up.hasValue(self.up.filter_file)) \
            or (self.up.hasValue(self.up.input2_file) and self.up.hasValue(self.up.filter_file)):
            Label(ws, text='', foreground='green').grid(row=4, columnspan=3, pady=10)
            pb1 = Progressbar(
                ws, 
                orient=HORIZONTAL,
                length=300, 
                mode='determinate'
                )
            pb1.grid(row=4, columnspan=3, pady=20)
            self.up.uploadFiles()
            for i in range(5):
                ws.update_idletasks()
                pb1['value'] += 20
                time.sleep(1)
            pb1.destroy()
            self.show_label( text="Excel Copied Successfully!", color="green")
        else:
            self.show_label( text="Please upload a valid file", color="red")

    def show_label_for_file_upload(self,text,color):
        lbl1=Label(self.ws, text=text, foreground=color)
        lbl1.grid(row=4, columnspan=3, pady=5)


    def show_label(self,text,color):
        self.show_label_for_file_upload("                                        ","white")
        Label(self.ws, text=text, foreground=color).grid(row=6, columnspan=3, pady=10)

ws = Tk()
app = Application(ws)
ws.mainloop()
