import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
    
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import *

import pandas as pd
import numpy as np

from PIL import ImageTk, Image

import sys
import os

from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import os

f               = None
selectedFile    = 'START.xlsx'
startDate       = '2014-4-16'
endDate         = '2014-4-17'
Seismograph     = 'None'
SerialNumber    = 'None'
Date            = 'None'
Client          = 'None'
ContactName     = 'None'
ContactAddress  = 'None'
ProjectName     = 'None'
ProjectNumber   = 'None'
ProjectDuration = 'None'
ProjectLocation = 'None'
allText = ''
    
class GeoViewerapp(tk.Tk):
    
    def __init__(self):
        
        tk.Tk.__init__(self)
        tk.Tk.title(self,'GeoViewer (Version 1.0)')
        tk.Tk.state(self,'normal')
        tk.Tk.configure(self,bg='white')
        
        #toolbar
        toolbar = tk.Frame(self, height=50, relief='raised', borderwidth=2)
        toolbar.pack(side='top', fill='x')
        #working window
        container = tk.Frame(self, bg='white', relief='sunken')
        container.pack(side='right', fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        #filemenu
        menubar = tk.Menu(container)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label='New Project', command=lambda: self.show_frame(PageOne))
        filemenu.add_separator()
        filemenu.add_command(label='Save', command = self.save)
        filemenu.add_separator()
        filemenu.add_command(label='Exit', command=Exit)
        menubar.add_cascade(label='File', menu=filemenu)
        
        #helpmenu
        helpMenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=helpMenu)
        helpMenu.add_command(label="About")
        
        tk.Tk.config(self, menu=menubar)

        topbar = tk.Frame(self, bg='white', height=50)
        topbar.pack(side='top', fill='x', pady=15, padx=160)
        
        button1 = ttk.Button(topbar, text="Select Seismograph", command=self.popup1)
        button1.pack(side='left', anchor='center')
        
        button2 = ttk.Button(topbar, text="Select Data", command=self.selectdata1)
        button2.pack(side='left', anchor='center')
        
        button3 = ttk.Button(topbar, text="Set Duration", command=self.popup)
        button3.pack(side='left', anchor='center')
        
        button4 = ttk.Button(topbar, text="Generate Report")
        button4.pack(side='left', anchor='center')
        
        button5 = ttk.Button(topbar, text="Enter Project Info", command=self.popup2)
        button5.pack(side='left', anchor='center')
        
        global f
        f = Figure(figsize=(8,20), dpi=80)
        # f.text(.9, .8,allText)
        f.suptitle('Peak Particle Velocity vs. Time', fontsize=20)
        a = f.add_subplot(1,1,1)
        a.set_xlabel('Time')
        a.set_ylabel('Peak Particle Velocity (in/s)')
        
        global selectedFile
        aData = pd.read_excel(selectedFile, sheet_name='Sheet1') #, nrows = customDuration)
        #vDeformation  = aData['Deformation']
        vPPVX         = aData['X']
        vPPVY         = aData['Y']
        vPPVZ         = aData['Z']
        vDates        = aData['Time']
         
        mask = (vDates > startDate) & (vDates <= endDate)
        aData = aData.loc[mask]
        #vDeformation  = aData['Deformation']
        vPPVX         = aData['X']
        vPPVY         = aData['Y']
        vPPVZ         = aData['Z']
        vDates        = aData['Time']

        a.plot(vDates,vPPVX)
        a.plot(vDates,vPPVY)
        a.plot(vDates,vPPVZ)
        
        canvas = FigureCanvasTkAgg(f, self)
        canvas.draw()
        canvas.get_tk_widget().pack(side='bottom', pady=20, expand=True, fill='both', anchor='s')

    def popup(self):
        self.w=popupWindow(self.master)
        self.wait_window(self.w.top)
        app.destroy()
        self.__init__()
    
    def popup1(self):
        self.w=SelectInstrument(self.master)
        self.wait_window(self.w.top)
        app.destroy()
        self.__init__()
        
    def popup2(self):
        self.w=ProjectInfo(self.master)
        self.wait_window(self.w.top)
        app.destroy()
        self.__init__()
    
    def save(self): 
        global f
        f.savefig("plot.pdf", bbox_inches='tight')

    def selectdata1(self): 
        global selectedFile
        selectedFile = filedialog.askopenfilename(initialdir = "/",title = "Select Data",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
        app.destroy()
        self.__init__()
        
        # data = pd.read_excel(C, sheet_name='Sheet1')
        # df = pd.DataFrame(data, columns=['Time','PPVX'])
        # vDates = df['Time']
        # vPPVX  = df['PPVX']       

def popupmsg():
    popup = tk.Toplevel()
    popup.wm_title("!")
    label = ttk.Label(popup, text="msg")
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()
    
def Exit():
    app.destroy()


#def save(): 
    #files = [('All Files', '*.*'),  
             #('Python Files', '*.py'), 
             #('Text Document', '*.txt')] 
    #file = filedialog.asksaveasfile(filetypes = files, defaultextension = '.txt') 
    #f.savefig('foo.pdf', bbox_inches='tight')

class popupWindow(object):
    def __init__(self,master):
        top=self.top=Toplevel(master)
        self.l=Label(top,text="Please input the start and end dates.")
        self.l.pack()
        self.a=Label(top,text="Format: yyyy-mm-dd") 
        self.a.pack()
        self.s=Entry(top, text='Start Date')
        self.s.pack()
        self.e=Entry(top)
        self.e.pack()
        self.b=Button(top,text='Ok',command=self.cleanup)
        self.b.pack()
    def cleanup(self):
        global startDate
        global endDate
        startDate = str(self.s.get())
        endDate = str(self.e.get())
        # self.value=self.e.get()
        self.top.destroy()

class SelectInstrument(object):
    def __init__(self,master):
        top=self.top=Toplevel(master)
        self.l=Label(top,text='Please input the Seismograph model and serial number.')
        self.l.grid(row=1,column=1,columnspan=2,pady=5)
        self.S=Label(top,text='Seismograph Model')
        self.S.grid(row=2,column=1)
        self.N=Label(top,text='Serial Number')
        self.N.grid(row=3,column=1)
        self.s=Entry(top, text='Seismograph Model')
        self.s.grid(row=2,column=2)
        self.e=Entry(top, text='Serial Number')
        self.e.grid(row=3,column=2)
        self.b=Button(top,text='Ok',command=self.cleanup)
        self.b.grid()
    def cleanup(self):
        global Seismograph
        global SerialNumber
        Seismograph = str(self.s.get())
        SerialNumber = str(self.e.get())
        # self.value=self.e.get()
        self.top.destroy()

class ProjectInfo(object):
    def __init__(self,master):
        top=self.top=Toplevel(master)
        top.geometry('415x280')
        self.l=Label(top,text='Please input the Project Details.')
        self.l.grid(row=1,column=1,columnspan=2,pady=5)
        self.D=Label(top,text='Date (yyyy-mm-dd)')
        self.D.grid(row=2,column=1)
        self.C=Label(top,text='Client')
        self.C.grid(row=3,column=1)
        self.N=Label(top,text='Contact Name')
        self.N.grid(row=4,column=1)
        self.A=Label(top,text='Contact Address')
        self.A.grid(row=5,column=1)
        self.P=Label(top,text='Project Name')
        self.P.grid(row=6,column=1)
        self.U=Label(top,text='Project Number')
        self.U.grid(row=7,column=1)
        self.T=Label(top,text='Project Duration (months)')
        self.T.grid(row=8,column=1)
        self.L=Label(top,text='Project Location')
        self.L.grid(row=9,column=1)
        #----------------------------------------------------------------#
        self.d=Entry(top, text='Date')
        self.d.grid(row=2,column=2)
        self.c=Entry(top, text='Client')
        self.c.grid(row=3,column=2)
        self.n=Entry(top, text='Contact Name')
        self.n.grid(row=4,column=2)
        self.a=Entry(top, text='Contact Address')
        self.a.grid(row=5,column=2)
        self.p=Entry(top, text='Project Name')
        self.p.grid(row=6,column=2)
        self.u=Entry(top, text='Project Number')
        self.u.grid(row=7,column=2)
        self.t=Entry(top, text='Project Duration')
        self.t.grid(row=8,column=2)
        self.l=Entry(top, text='Project Location')
        self.l.grid(row=9,column=2)
        #----------------------------------------------------------------#
        self.b=Button(top,text='Ok',command=self.cleanup)
        self.b.grid()
    def cleanup(self):
        global Date
        global Client
        global ContactName
        global ContactAddress
        global ProjectName
        global ProjectNumber
        global ProjectDuration
        global ProjectLocation
        global allText
        Date = str(self.d.get())
        Client = str(self.c.get())
        ContactName = str(self.n.get())
        ContactAddress = str(self.a.get())
        ProjectName = str(self.p.get())
        ProjectNumber = str(self.u.get())
        ProjectDuration = str(self.t.get())
        ProjectLocation = str(self.l.get())
        allText += ("Date: " + Date + '\n')
        allText += ("Client: " + Client + '\n')
        allText += ("Contact Name: " + ContactName + '\n')
        allText += ("Contact Address: " + ContactAddress + '\n')
        allText += ("Project Name: " + ProjectName + '\n')
        allText += ("Project Number: " + ProjectNumber + '\n')
        allText += ("Project Duration: " + ProjectDuration + '\n')
        allText += ("Project Location: " + ProjectLocation + '\n')

        pdf = FPDF()
        pdf.add_page()
        pdf.set_xy(0, 0)
        pdf.set_font('arial', 'B', 13.0)
        pdf.multi_cell(0, 5, allText)
        # pdf.cell(ln=10, h=10.0, align='L', w=0, txt=allText, border=0)
        pdf.output('test.pdf', 'F')

        global f
        f.subplots_adjust(left=0.3, right=0.9, bottom=0.3, top=0.9)
        #-------------Keeps the plot from being cut off-----------#
        f.savefig("plot.pdf", bbox_inches='tight')
        
        file1 = PdfFileReader(open("test.pdf", "rb"))
        file2 = PdfFileReader(open("plot.pdf", "rb"))

        page = file1.getPage(0)
        page.mergePage(file2.getPage(0))

        output = PdfFileWriter()
        output.addPage(page)
        outputStream = open("output.pdf", "wb")
        output.write(outputStream)
        outputStream.close()

        # self.value=self.e.get()
        self.top.destroy()
        

app = GeoViewerapp()
app.configure(background='white')
app.geometry('900x900')
# app.iconbitmap(r'icon.ico')
app.mainloop()