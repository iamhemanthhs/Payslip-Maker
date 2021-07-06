import random
import time
import csv
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import tkinter as tk

from tkinter import *
from tkinter import Toplevel,messagebox,filedialog
import requests
from tkinter import *
from tkinter import Toplevel,messagebox,filedialog
from tkinter.ttk import Treeview
from tkinter import ttk


root = Tk()
root.title('Team Tree Enterprises')
root.config(bg='#0e1621')
root.geometry('480x100+700+400')
root.iconbitmap('capture.ico')
root.resizable(False,False)

Frame1=Frame(root,bg='#243142')
Frame1.place(x=30,y=30,width=80,height=40)

Frame2=Frame(root,bg='#243142')
Frame2.place(x=140,y=30,width=70,height=40)

Frame3=Frame(root,bg='#243142')
Frame3.place(x=240,y=30,width=70,height=40)

Frame4=Frame(root,bg='#243142')
Frame4.place(x=340,y=30,width=110,height=40)


def template():
    global ff
    ff = filedialog.askopenfilename()
    return ff

def data():
    global cc
    cc = filedialog.askopenfilename()
    return cc

def result():
    global pdffiles
    pdffiles = filedialog.askdirectory()
    return pdffiles

def payslip():
    csvfn = r"{}".format(cc)

    def mkw(n):
        tpl1 = r"{}".format(ff)
        tpl = DocxTemplate(tpl1)
        print(tpl)

        filepath = r'{}'.format(pdffiles)

        # tpl = DocxTemplate("template.docx") # In same directory
        df = pd.read_csv(csvfn)
        df_to_doct = df.to_dict()  # dataframe -> dict for the template render
        x = df.to_dict(orient='records')
        context = x
        tpl.render(context[n])
        tpl.save("{}/%s.docx".format(filepath) % str(n + 1))
        wait = time.sleep(random.randint(1, 2))

    df2 = len(pd.read_csv(csvfn))
    print("There will be ", df2, "files")

    for i in range(0, df2):
        print("Making file: ", f"{i},", "..Please Wait...")
        mkw(i)

    print("Done! - Now check your files")

    newpath = r'{}'.format(pdffiles)
    print(newpath)
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    i = df2
    for a in range(1, i + 1):
        convert("{}/{}.docx".format(newpath,a), r"{}/".format(pdffiles))
    messagebox.showinfo('Notification', 'Sucussfully Done.....')

    dbl1 = r"{}".format(pdffiles)
    x=df2
    for x in range(1,x+1):
        os.remove("{}/{}.docx".format(newpath,x))
    messagebox.showinfo('Notification','Successfully Deleted all word files...Payslips are ready!')



tmpbtn=Button(Frame1,text='Template',font=('Italic',11,NORMAL),bd=6,bg='#17212b',fg='white',command=template)
tmpbtn.place(x=0,y=0,width=80,height=40)


databtn=Button(Frame2,text='Data',font=('Italic',11,NORMAL),bd=6,bg='#17212b',fg='white',command=data)
databtn.place(x=0,y=0,width=70,height=40)

rsltbtn=Button(Frame3,text='Result',font=('Italic',11,NORMAL),bd=6,bg='#17212b',fg='white',command=result)
rsltbtn.place(x=0,y=0,width=70,height=40)

pslpbtn=Button(Frame4,text='Make Payslip',font=('Italic',11,NORMAL),bd=6,bg='#17212b',fg='white',command=payslip)
pslpbtn.place(x=0,y=0,width=110,height=40)


root.mainloop()






