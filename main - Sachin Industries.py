#drag and drop enabled
# import all components
#button(..., image = resetimg)
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import asksaveasfile
from tkinter import filedialog
import TkinterDnD2
from TkinterDnD2 import *

import pandas as pd

import xlsxwriter as xl
from win32com import client

import math
import os

import datetime
from datetime import date

T = None
c = None
pp = None
m = None
sw = None
filepath = ''
filename = ''
verifyur = 0
x = ''
folderpath=''
last = ''
material = ''
tempcname = ''
try:
        dfdv = pd.read_csv('Default_Values.csv')
        dfssal = pd.read_csv('SS_AL_Feed_Rate.csv')
        dfmscu = pd.read_csv('MS_CU_Feed_Rate.csv')
except:
        messagebox.showerror("Error", "Some files are missing!")
        exit()
mnths=['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
# Function for opening the file explorer root
def browseFiles():
        global filepath
        global filename
        global folderpath
        global last
        global cname
        global tempcname
        last = 'file'
        filepath = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = (("Hnf files","*.hnf*"),("all files","*.*")))
        if filepath != '':
                filename = filepath[(filepath.rindex('/')+1):]
                # Change label contents
                label_file_explorer.configure(text="File Opened: "+filepath)
                text12.config(text = "Reading File: " + filename)
                folderpath = filepath[:filepath.rindex('/')]
                '''if filepath.lower().find('desktop') == -1:
                        tempcname = filepath[0:filepath.rindex('/')][filepath[0:filepath.rindex('/')].rindex('/')+1:]
                else:
                        cnt = 0
                        for i in mnths:
                                if filepath.lower().find(i) == -1:
                                        cnt = 1
                                        m = i
                        if cnt == 1:
                                tempcname = (filepath[filepath.lower().find('september')+1:][filepath[filepath.lower().find('september'):].find('/'):])[:(filepath[filepath.lower().find('september')+1:][filepath[filepath.lower().find('september'):].find('/'):]).find('/')]
                        else:
                                tempcname = filepath[filepath.find('/Desktop')+1:][filepath[filepath.find('/Desktop')+1:].find('/')+1:][:filepath[filepath.find('/Desktop')+1:][filepath[filepath.find('/Desktop')+1:].find('/')+1:].find('/')]
                cname.delete(0, END)
                cname.insert(0, tempcname)
                print('tempcname=',tempcname)'''

def browseFolder():
        global foldername
        global folderpath
        global last
        global cname
        global tempcname
        last = 'folder'
        folderpath = filedialog.askdirectory()
        
        try:
                foldername = folderpath[(folderpath.rindex('/')+1):]
                # Change label contents
                label_file_explorer.configure(text="Folder Opened: "+folderpath)
                text12.config(text = "Reading Folder: " + foldername)
                '''if folderpath.lower().find('desktop') == -1:
                        tempcanme = folderpath[folderpath.rindex('/')+1:]
                else:
                        cnt = 0
                        for i in mnths:
                                if folderpath.lower().find(i) == -1:
                                        cnt = 1
                                        m = i
                        if cnt == 1:
                                tempcname = (folderpath[folderpath.lower().find('september')+1:][folderpath[folderpath.lower().find('september'):].find('/'):])[:(folderpath[folderpath.lower().find('september')+1:][folderpath[folderpath.lower().find('september'):].find('/'):]).find('/')]
                        else:
                                tempcname = folderpath[folderpath.find('/Desktop')+1:][folderpath[folderpath.find('/Desktop')+1:].find('/')+1:][:folderpath[folderpath.find('/Desktop')+1:][folderpath[folderpath.find('/Desktop')+1:].find('/')+1:].find('/')]
                cname.delete(0, END)
                cname.insert(0, tempcname)
                print(tempcname)'''
        except:
                pass
                
# Create the root root
#root = Tk()
root = TkinterDnD.Tk()

#resetimage = 
#resetimg = PhotoImage(file = r'reset.png').subsample(30,30)
settingimg = PhotoImage(file = r'setting1.png').subsample(1,1)

# Set root title
root.title('Sachin Industries')
Grid.columnconfigure(root, 0, weight=1)
Grid.columnconfigure(root, 1, weight=1)

# Set root size
root.geometry("1500x1000")
root.state('zoomed')

#table
# columns
columns = ('#1', '#2', '#3', '#4', '#5', '#6', '#7', '#8')

style = ttk.Style()
style.theme_use('classic')

tree = ttk.Treeview(root, columns=columns, show='headings')
tree.column("#1", anchor = 'center')
tree.column("#2", anchor = 'center')
tree.column("#3", anchor = 'center')
tree.column("#4", anchor = 'center')
tree.column("#5", anchor = 'center')
tree.column("#6", anchor = 'center')
tree.column("#7", anchor = 'center')
tree.column("#8", anchor = 'center', width = -10)

# define headings
tree.heading('#1', text='File Name')
tree.heading('#2', text='Labor/With Material')
tree.heading('#3', text='Material')
tree.heading('#4', text='Thickness')
tree.heading('#5', text='Quantity')
tree.heading('#6', text='Per Peice Cost')
tree.heading('#7', text='Total Cost')
tree.heading('#8', text='')
# add a scrollbar
scrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
#table ends

def feedratedef(m,thick):
        fr=0
        if m == 'SS' or m == 'AL':
                for i in range(1,len(dfssal['thickness <='])):
                        if float(dfssal['thickness <='][i]) <= thick:
                                fr = dfssal['feed rate'][i]
                                ind = i

        elif m == 'MS' or m == 'CU' :
                for i in range(1,len(dfmscu['thickness <='])):
                        if float(dfmscu['thickness <='][i]) <= thick:
                                fr = dfmscu['feed rate'][i]
                                ind = i
        return fr

def valuesfromfile():
        global filename
        global filepath
        global T
        global c
        global pp
        global m
        global sw
        filename = filepath[(filepath.rindex('/')+1):]
        data = open(filepath,"r+").read()
        T = float(data[data.find("T= ")+3:][0:data[data.find("T= ")+3:].find(' ')])
        c = float(data[data.find("Cutting way: ")+13:][0:data[data.find("Cutting way: ")+13:].find(' ')])
        pp = int(data[data.find("Pierce Qty: ")+12:][0:data[data.find("Pierce Qty: ")+12:].find(' ')])
        m = data[data.find("Material: ")+10:][0:data[data.find("Material: ")+10:].find(' ')]
        sw = float(data[data.find("Sheet Weight: ")+14:][0:data[data.find("Sheet Weight: ")+14:].find(' ')])

        if m == 'mild':
                m = 'MS'
        elif m == 'stainless':
                m = 'SS'
        elif m == 'aluminium':
                m = 'AL'
        elif m == 'copper':
                m = 'CU'
        algoauto()

#algoauto
def algoauto():
        global verifyur
        
        global material
        global filename
        global filepath

        global T
        global c
        global pp
        global m
        global sw
        print('2nd filepath =', filepath)
        '''
        filename = filepath[(filepath.rindex('/')+1):]
        data = open(filepath,"r+").read()
        T = float(data[data.find("T= ")+3:][0:data[data.find("T= ")+3:].find(' ')])
        c = float(data[data.find("Cutting way: ")+13:][0:data[data.find("Cutting way: ")+13:].find(' ')])
        pp = int(data[data.find("Pierce Qty: ")+12:][0:data[data.find("Pierce Qty: ")+12:].find(' ')])
        m = data[data.find("Material: ")+10:][0:data[data.find("Material: ")+10:].find(' ')]
        sw = float(data[data.find("Sheet Weight: ")+14:][0:data[data.find("Sheet Weight: ")+14:].find(' ')])

        if m == 'mild':
                m = 'MS'
        elif m == 'stainless':
                m = 'SS'
        elif m == 'aluminium':
                m = 'AL'
        elif m == 'copper':
                m = 'CU'
        '''


        fr = feedratedef(m, T)


        mscr1 = mscr.get()
        if m == 'AL':
                alcr1 = int(dfdv['Values'][2])
        if m == 'CU':
                cucr1 = int(dfdv['Values'][3])
        if mscr1 == '':
                mscr1 = int(dfdv['Values'][0])
        else:
                mscr1 = int(mscr1)
        sscr1 = sscr.get()
        if sscr1 == '':
                sscr1 = int(dfdv['Values'][1])
        else:
                sscr1 = int(sscr1)
        
        if m == 'MS':
            time = c/fr
            cost = math.ceil((time*mscr1)+(T*pp))
        elif m == 'SS':
            time = c/fr
            cost = math.ceil((time*sscr1)+(T*pp))
        elif m == 'AL':
            time = c/fr
            cost = math.ceil((time*alcr1)+(T*pp))
        elif m == 'CU':
            time = c/fr
            cost = math.ceil((time*cucr1)+(T*pp))

            
        msr1 = msr.get()
        if msr1 == '':
                msr1 = int(dfdv['Values'][4])
        else:
                mscr1 = int(msr1)
        ssr1 = ssr.get()
        if ssr1 == '':
                ssr1 = int(dfdv['Values'][5])
        else:
                ssr1 = int(ssr1)

                
        if m == 'MS':
            if var1.get() == 1:
                mc = int(dfdv['Values'][4])
                if ssr == '':
                        mc = int(dfdv['Values'][4])
                else:
                        mc = int(msr.get())
                cost = cost + (mc*sw)               
        elif m == 'SS':
            if var2.get() == 1:
                mc = int(dfdv['Values'][5])
                if ssr == '':
                        mc = int(dfdv['Values'][5])
                else:
                        mc = int(ssr.get())
                cost = cost + (mc*sw)
        if qty.get() == '':
                q = int(dfdv['Values'][6])
        else:
                q = int(qty.get())
        if q<=0:
            q = int(dfdv['Values'][6])
        ppc = cost
        cost = cost *q
        print('Time =', time)
        print('Cost =', cost)
        mmaterial = ''
        smaterial = ''
        cost = float(math.ceil(cost))
        print("verifyur = ",verifyur)
        if verifyur == 1:
                ls = [filename, '', m, T, q, ppc, cost, filepath]
                return ls
        if m == 'MS':
                if var1.get() == 1:
                        mmaterial = 'With Material'
                else:
                        mmaterial = 'Labour Only'
                tree.insert('', 'end', text="1", values=(filename, mmaterial, m, T, q, ppc, cost, filepath))
        elif m == 'SS':
                if var2.get() == 1:
                        smaterial = 'With Material'
                else:
                        smaterial = 'Labour Only'
                tree.insert('', 'end', text="1", values=(filename, smaterial, m, T, q, ppc, cost, filepath))
        elif m == 'CU' or 'AL':
                tree.insert('', 'end', text="1", values=(filename, 'Labour Only', m, T, q, ppc, cost, filepath))
        tamt.config(text = (float(tamt.cget('text'))+cost))


def folder():
        global folderpath
        global filepath
        os.chdir(folderpath)
        for file in os.listdir():
                # Check whether file is in text format or not
                if file.endswith(".Hnf"):
                        if folderpath[-1] == '/':
                                filepath = f"{folderpath}{file}"
                        else:
                                filepath = f"{folderpath}/{file}"
                        valuesfromfile()

def cnamechange():
        if folderpath.find('/') != -1:
                if folderpath.lower().find('desktop') == -1:
                        tempcname = folderpath[folderpath.rindex('/')+1:]
                else:
                        cnt = 0
                        for i in mnths:
                                if folderpath.lower().find(i) != -1:
                                        cnt = 1
                                        m = i
                        if cnt == 1:
                                tempcname = (folderpath[folderpath.lower().find(m)+1:][folderpath[folderpath.lower().find(m):].find('/'):])[:(folderpath[folderpath.lower().find(m)+1:][folderpath[folderpath.lower().find(m):].find('/'):]).find('/')]
                        else:
                                tempcname = folderpath[folderpath.find('/Desktop')+1:][folderpath[folderpath.find('/Desktop')+1:].find('/')+1:][:folderpath[folderpath.find('/Desktop')+1:][folderpath[folderpath.find('/Desktop')+1:].find('/')+1:].find('/')]
                cname.delete(0, END)
                cname.insert(0, tempcname)

def go():
        global folderpath
        global filepath
        global mnths
        m=''
        
        cnamechange()
        
        if last == '':
                messagebox.showerror("Error", "Please browse for a file/folder")
        elif filepath == '' and folderpath == '':
                messagebox.showerror("Error", "Please browse and select a file/folder properly")
        elif filepath != '' and last == 'file':
                valuesfromfile()
        elif folderpath != '' and last == 'folder':
                folder()
def exportxl():
        global last
        global foldername
        global filepath
        if tree.get_children() == ():
                messagebox.showerror("Error", "No data entered")
                return
        fxm = [('Excel', '*.xlsx')]
        fx = None
        iname = None
        print(cname.get())
        print(type(cname.get()))
        #print(filepath[(filepath[0: filepath.rindex('/')].rindex('/'))+1: filepath.rindex('/')])
        if len(cname.get()) != 0:
                print('in cname')
                iname = cname.get()
        elif last == 'folder' and foldername != '':
                iname = foldername
        elif last == 'file' and filepath != '':
                print('hi 1')
                if filepath.count('/') > 1:
                        iname = filepath[(filepath[0: filepath.rindex('/')].rindex('/'))+1: filepath.rindex('/')]
                else:
                        iname = filepath[0: filepath.rindex('/')]
        print('last =', last)
        print('iname =',iname)
        try:
                fx = asksaveasfile(filetypes = fxm, defaultextension = fxm, initialfile = 'Costing sheet')
        except:
                messagebox.showerror("Error", "First close the excel which is already open, then try again")
        if fx != None:
                workbook = xl.Workbook(fx.name)
                worksheet = workbook.add_worksheet()
                print("Hello 1")
                cellf1 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'font_size' : 18, 'font':'Bell MT', 'font_color' : '#ff0000'})
                cellf2 = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'font_size' : 18, 'font':'Bell MT', 'font_color' : '#384a9c'})
                cellf3 = workbook.add_format({'valign': 'vcenter', 'border': 1, 'bold': True})
                cellf4 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                cellf5 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold':True})
                
                worksheet.set_row(0, 30)
                worksheet.set_row(1, 30)
                worksheet.set_row(2, 30)
                worksheet.set_row(3, 5)
                worksheet.set_row(4, 27)
                worksheet.set_row(5, 23)
                worksheet.set_row(6, 5)
                worksheet.set_row(7, 20)
                worksheet.set_row(8, 20)
                worksheet.set_row(9, 5)
                worksheet.set_row(10, 30)
                
                worksheet.set_column(0, 0, 35)
                worksheet.set_column(1, 1, 15)
                worksheet.set_column(2, 2, 10)
                worksheet.set_column(3, 3, 10)
                worksheet.set_column(4, 4, 17)
                worksheet.set_column(5, 5, 23)
                worksheet.set_column(6, 6, 23)

                worksheet.merge_range('A1:G1', 'CASH MEMO', cellf1)
                worksheet.merge_range('A2:G2', 'Sachin Technologies', cellf2)
                worksheet.merge_range('A3:G3', 'S-67, MIDC Bhosari, Pin 411026', cellf2)

                worksheet.write('A5', 'Company Name: ', cellf3)
                worksheet.merge_range('B5:E5', cname.get(), cellf3)
                worksheet.write('A6', 'Address: ', cellf3)
                worksheet.merge_range('B6:G6', '', cellf3)
                print("Hello 2")
                worksheet.write('F5', 'Date', cellf3)
                today = date.today()
                mnth=''
                if today.month == 1:
                        mnth = 'Jan'
                elif today.month == 2:
                        mnth = 'Feb'
                elif today.month == 3:
                        mnth = 'Mar'
                elif today.month == 4:
                        mnth = 'Apr'
                elif today.month == 5:
                        mnth = 'May'
                elif today.month == 6:
                        mnth = 'June'
                elif today.month == 7:
                        mnth = 'July'
                elif today.month == 8:
                        mnth = 'Aug'
                elif today.month == 9:
                        mnth = 'Sep'
                elif today.month == 10:
                        mnth = 'Oct'
                elif today.month == 11:
                        mnth = 'Nov'
                elif today.month == 12:
                        mnth = 'Dec'
                
                worksheet.write('G5', str(today.day)+' '+mnth+' '+str(today.year), cellf3)

                worksheet.merge_range('A8:B8', 'Only Laser Cutting Charges', cellf3)
                worksheet.merge_range('A9:B9', 'Laser Cutting with Material Charges', cellf3)
                worksheet.write('C8', '', cellf3)
                worksheet.write('C9', '', cellf3)
                
                worksheet.write('A11', 'Particular', cellf4)
                worksheet.write('B11', 'Material', cellf4)
                worksheet.write('C11', 'Thickness', cellf4)
                worksheet.write('D11', 'Gas', cellf4)
                #worksheet.write('E11', 'per piece Rate(only Laser)', cellf5)
                worksheet.write('E11', 'Quantity', cellf4)
                worksheet.write('F11', 'Amount', cellf4)
                worksheet.write('G11', 'Total Amount', cellf5)
                print("Hello 3")
                wmcnt = 0
                olcnt = 0
                errorcnt = 0
                gas = ''
                sum = 0
                i = 12
                for a in tree.get_children():
                        m = tree.item(a, 'values')[2]
                        t = float(tree.item(a, 'values')[3])
                        if m == 'MS':
                                if t <= 2.0:
                                        gas = 'N2'
                                else:
                                        gas = 'O2'
                        elif m == 'SS' or m == 'AL':
                                gas = 'N2'
                        elif m == 'CU':
                                if t <= 3.0:
                                        gas = 'N2'
                                else:
                                        gas = 'O2'
                        if gas == '':
                                print('check this out !!!!!!!!!!!!!!', t)

                        worksheet.write('A'+str(i), (tree.item(a, 'values')[0])[0:(tree.item(a, 'values')[0].find('.hnf'))], cellf4)
                        worksheet.write('B'+str(i), tree.item(a, 'values')[2], cellf4)
                        worksheet.write('C'+str(i), tree.item(a, 'values')[3]+' mm', cellf4)
                        worksheet.write('D'+str(i), gas, cellf4)
                        worksheet.write('E'+str(i), int(tree.item(a, 'values')[4]), cellf4)
                        worksheet.write('F'+str(i), float(tree.item(a, 'values')[5]), cellf4)
                        worksheet.write('G'+str(i), ('=PRODUCT(E'+str(i)+',F'+str(i)+')'), cellf4)
                        if tree.item(a, 'values')[1] == 'With Material':
                                wmcnt+=1
                        elif tree.item(a, 'values')[1] == 'Labour Only':
                                olcnt+=1
                        else:
                                errorcnt+=1
                        #total cost
                        sum += float(tree.item(a, 'values')[5])
                        i+=1
                worksheet.write('E'+str(i), 'Total', cellf4)
                worksheet.write('F'+str(i), '=SUM(F12:F'+str(i-1)+')', cellf5)
                worksheet.write('G'+str(i), '=SUM(G12:G'+str(i-1)+')', cellf5)
                if errorcnt != 0:
                        pass
                elif wmcnt != 0 and olcnt != 0:
                        pass
                elif wmcnt == i-12:
                        worksheet.write('C9', 'Yes', cellf3)
                elif olcnt == i-12:
                        worksheet.write('C8', 'Yes', cellf3)
                print("Hello 4")
                print(errorcnt)
                print(wmcnt)
                print(olcnt)
                print(i, i-12)
                workbook.close()
                '''try:
                        workbook.close()
                except:
                        messagebox.showerror("Error", "First close the excel which is already open, then try again")'''
        else:
                pass
        fx = None

def only_numbers(char):
        return char.isdigit()

def only_float(s, S1, d, i):
        i = int(i)
        if d == '1':
                S = S1+s
        elif d == '0':
                S = S1.replace(s, '', i-1)
        else:
                return True
        try:
                float(S)
        except:
                if S1 == '.' and d == '0':
                        return True
                if S == '.' and d == '0':
                        return True
                if S == '' and d == '0':
                        return True
                if s == '.':
                        cnt = 0
                        for i in S:
                                if i == '.':
                                        cnt+= 1
                        if d == '1' and cnt == 0:
                                return True
                        if d == '0' and cnt >= 1:
                                return True
                return False
        return True

def test(char):
        pass


def mcalcu():
        def test123():
                global T
                global c
                global pp
                global m
                global sw
                T = float(thickm.get())
                c = float(cpm.get())
                pp = int(ppm.get())
                m = mm.get()
                sw = float(swm.get())
                filename = particularm.get()
                fr = feedratedef(m, T)


                mscr1 = mscr.get()
                if m == 'AL':
                        alcr1 = 200
                if m == 'CU':
                        cucr1 = 150
                if mscr1 == '':
                        mscr1 = 85
                else:
                        mscr1 = int(mscr1)
                sscr1 = sscr.get()
                if sscr1 == '':
                        sscr1 = 125
                else:
                        sscr1 = int(sscr1)
                
                if m == 'MS':
                    time = c/fr
                    cost = math.ceil((time*mscr1)+(T*pp))
                elif m == 'SS':
                    time = c/fr
                    cost = math.ceil((time*sscr1)+(T*pp))
                elif m == 'AL':
                    time = c/fr
                    cost = math.ceil((time*alcr1)+(T*pp))
                elif m == 'CU':
                    time = c/fr
                    cost = math.ceil((time*cucr1)+(T*pp))

                    
                msr1 = msr.get()
                if msr1 == '':
                        msr1 = 95
                else:
                        mscr1 = int(msr1)
                ssr1 = ssr.get()
                if ssr1 == '':
                        ssr1 = 120
                else:
                        ssr1 = int(ssr1)

                        
                if m == 'MS':
                    if var1.get() == 1:
                        mc = 95
                        if ssr == '':
                                mc = 95
                        else:
                                mc = int(msr.get())
                        cost = cost + (mc*sw)               
                elif m == 'SS':
                    if var2.get() == 1:
                        mc = 120
                        if ssr == '':
                                mc = 120
                        else:
                                mc = int(ssr.get())
                        cost = cost + (mc*sw)
                if qty.get() == '':
                        q = 1
                else:
                        q = int(qty.get())
                if q<=0:
                    q = 1
                cost = cost *q
                print('Time =', time)
                print('Cost =', cost)
                mmaterial = ''
                smaterial = ''
                cost = float(math.ceil(cost))
                '''
                if m == 'MS':
                        if var1.get() == 1:
                                mmaterial = 'With Material'
                        else:
                                mmaterial = 'Labour Only'
                        tree.insert('', 'end', text="1", values=(filename, mmaterial, material, T, q, cost))
                elif m == 'SS':
                        if var2.get() == 1:
                                smaterial = 'With Material'
                        else:
                                smaterial = 'Labour Only'
                        tree.insert('', 'end', text="1", values=(filename, smaterial, material, T, q, cost))
                elif m == 'CU' or 'AL':
                        tree.insert('', 'end', text="1", values=(filename, 'Labour Only', material, T, q, cost))
                tamt.config(text = (float(tamt.cget('text'))+cost))'''

        def autocalcpprm(char):
                if thickm.get() == '.':
                        #thickm.configure(validatecommand=None)
                        thickm.delete(0, END)
                        thickm.insert(0, '0.0')
                        #thickm.configure(validatecommand=(root.register(only_float), '%S'))
                if thickm.get != '' and cpm.get() != '' and ppm.get() != '':
                        T = float(thickm.get())
                        c = float(cpm.get())
                        pp = int(ppm.get())
                        m = mm.get()
                        fr = feedratedef(m, T)
                        mscr1 = mscr.get()
                        if m == 'AL':
                                alcr1 = 200
                        if m == 'CU':
                                cucr1 = 150
                        if mscr1 == '':
                                mscr1 = 85
                        else:
                                mscr1 = int(mscr1)
                        sscr1 = sscr.get()
                        if sscr1 == '':
                                sscr1 = 125
                        else:
                                sscr1 = int(sscr1)
                        
                        if m == 'MS':
                            time = c/fr
                            cost = math.ceil((time*mscr1)+(T*pp))
                        elif m == 'SS':
                            time = c/fr
                            cost = math.ceil((time*sscr1)+(T*pp))
                        elif m == 'AL':
                            time = c/fr
                            cost = math.ceil((time*alcr1)+(T*pp))
                        elif m == 'CU':
                            time = c/fr
                            cost = math.ceil((time*cucr1)+(T*pp))
                        print(cost)
        def autocalctrwmm():
                pass
        def autocalctamtm():
                pass
        win = Toplevel()
        win.title("Manual Calulator")
        win.columnconfigure(0, weight=1)
        win.columnconfigure(1, weight=1)
        win.geometry("1500x500")
        label = Label(win, text = "Manual Calulator", height = 4, fg = "blue", font=300)
        label.grid(column = 0, row = 0, sticky="nsew", columnspan=20)
        
        cnamemtext = Label(win, text="Enter Company Name", font=150)
        cnamemtext.grid(column = 0, row = 1, sticky="e")
        cnamem = Entry(win, borderwidth=5, bg='white', font=10)
        cnamem.grid(column = 1, row = 1, sticky="w")
        
        crmtext = Label(win, text="Enter Cutting Rate", font=150)
        crmtext.grid(column = 0, row = 2, sticky="e")
        crm = Entry(win, borderwidth=5, bg='white', font=10)
        crm.grid(column = 1, row = 2, sticky="w")
        
        mrmtext = Label(win, text="Enter material Rate", font=150)
        mrmtext.grid(column = 0, row = 3, sticky="e")
        mrm = Entry(win, borderwidth=5, bg='white', font=10)
        mrm.grid(column = 1, row = 3, sticky="w")

        '''*************************************************************************************************************************'''
        
        frame2m = Frame(win)
        frame2m.columnconfigure(0, weight=1)
        frame2m.columnconfigure(1, weight=1)
        frame2m.columnconfigure(2, weight=1)
        frame2m.columnconfigure(3, weight=1)

        textmscrm = Label(frame2m, text="MSCR")
        mscrm = Entry(frame2m, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
        textsscrm = Label(frame2m, text="SSCR")
        sscrm = Entry(frame2m, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
        textalcrm = Label(frame2m, text="Aluminium CR")
        alcrm = Entry(frame2m, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
        textcucrm = Label(frame2m, text="Copper CR")
        cucrm = Entry(frame2m, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')

        var1m = IntVar()
        var2m = IntVar()
        tick1m = Checkbutton(win, text="MS with Material", variable=var1m, cursor='target')
        tick2m = Checkbutton(win, text="SS with Material", variable=var2m, cursor='target')

        textmsrm = Label(win, text="MSR")
        msrm = Entry(win, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
        textssrm = Label(win, text="SSR")
        ssrm = Entry(win, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')

        frame2m.grid(column=0, row=4, columnspan=15, sticky="ns")
        textmscrm.grid(column = 0, row = 0)
        textsscrm.grid(column = 1, row = 0)
        textalcrm.grid(column = 2, row = 0)
        textcucrm.grid(column = 3, row = 0)
        mscrm.grid(column = 0, row = 1)
        sscrm.grid(column = 1, row = 1)
        alcrm.grid(column = 2, row = 1)
        cucrm.grid(column = 3, row = 1)

        mscrm.insert(0, dfdv['Values'][0])
        sscrm.insert(0, dfdv['Values'][1])
        alcrm.insert(0, dfdv['Values'][2])
        cucrm.insert(0, dfdv['Values'][3])

        tick1m.grid(column = 0, row = 5, columnspan = 5)
        tick2m.grid(column = 0, row = 6, columnspan = 5)

        textmsrm.grid(column = 0, row = 7, sticky="e")
        msrm.grid(column = 1, row = 7, sticky="w")
        msrm.insert(0, dfdv['Values'][4])
        textssrm.grid(column = 0, row = 8, sticky="e")
        ssrm.grid(column = 1, row = 8, sticky="w")
        ssrm.insert(0, dfdv['Values'][5])


        '''*************************************************************************************************************************'''
        framem = Frame(win)
        
        framem.columnconfigure(0, weight=1)
        framem.columnconfigure(1, weight=1)
        framem.columnconfigure(2, weight=1)
        framem.columnconfigure(3, weight=1)
        framem.columnconfigure(4, weight=1)
        framem.columnconfigure(5, weight=1)
        framem.columnconfigure(6, weight=1)
        framem.columnconfigure(7, weight=1)
        framem.columnconfigure(8, weight=1)
        framem.columnconfigure(9, weight=1)

        framem.grid(column = 0, row = 10, sticky="nsew", columnspan=20)

        particularmtext = Label(framem,text ='Particular')
        particularmtext.grid(row = 0, column = 0)
        particularm = Entry(framem, borderwidth=5, bg='white', font=10)
        particularm.grid(row = 1, column = 0)

        mtextm = Label(framem,text ='Material*').grid(row = 0, column = 1)
        mm = ttk.Combobox(framem, width = 27, textvariable = tk.StringVar(), state= "readonly")
        mm['values']=('MS','SS','CU','AL')
        mm.grid(row = 1, column = 1)
        mm.current(0)


        def onlytry(d, i, p, s, S, v, V):
                #print('char =', char)
                print('d =', d)
                print('i =', i)
                print('p =', p)
                print('s =', s)
                print('S =', S)
                print('v =', v)
                print('V =', V)
                print('W =', W)
                print()
                return True
        thickmtext = Label(framem,text ='Thickness*')
        thickmtext.grid(row = 0, column = 2)
        thickm = Entry(framem, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
        thickm.grid(row = 1, column = 2)
        thickm.insert(0, '0.0')

        cpmtext = Label(framem,text ='Cutting Parameter*')
        cpmtext.grid(row = 0, column = 3)
        cpm = Entry(framem, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
        cpm.grid(row = 1, column = 3)

        ppmtext = Label(framem,text ='Peircing Point*')
        ppmtext.grid(row = 0, column = 4)
        ppm = Entry(framem, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
        ppm.grid(row = 1, column = 4)

        pprmtext = Label(framem,text ='Per Peice Rate')
        pprmtext.grid(row = 0, column = 5)
        pprm = Entry(framem, borderwidth=5, bg='#e6e6e6', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
        pprm.grid(row = 1, column = 5)
        pprmbtn = Button(framem, text='Auto Calc', command = autocalcpprm)
        pprmbtn.grid(row = 2, column = 5, sticky='nsew')

        swmtext = Label(framem,text ='Sheet Weight')
        swmtext.grid(row = 0, column = 6)
        swm = Entry(framem, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
        swm.grid(row = 1, column = 6)

        rwmmtext = Label(framem,text ='Rate With Material')
        rwmmtext.grid(row = 0, column = 7)
        rwmm = Entry(framem, borderwidth=5, bg='#e6e6e6', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
        rwmm.grid(row = 1, column = 7)
        rwmmbtn = Button(framem, text='Auto Calc', command = autocalctrwmm)
        rwmmbtn.grid(row = 2, column = 7, sticky='nsew')

        qtymtext = Label(framem,text ='Quantity')
        qtymtext.grid(row = 0, column = 8)
        qtym = Entry(framem, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
        qtym.grid(row = 1, column = 8)
        qtym.insert(0, '1')

        tamtmtext = Label(framem,text ='Total Amount')
        tamtmtext.grid(row = 0, column = 9)
        tamtm = Entry(framem, borderwidth=5, bg='#e6e6e6', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
        tamtm.grid(row = 1, column = 9)
        tamtmbtn = Button(framem, text='Auto Calc', command = autocalctamtm)
        tamtmbtn.grid(row = 2, column = 9, sticky='nsew')

        testbtnm = Button(win, text='Test', command = test123)
        testbtnm.grid(row = 10, column = 9)
        def test2(char):
                particularm.delete(0, END)
                particularm.insert(0, 'Enter')
        def test3(char):
                particularm.delete(0, END)
                particularm.insert(0, 'Exit')
        mm.bind('<FocusOut>', autocalcpprm)
        thickm.bind('<FocusOut>', autocalcpprm)
        cpm.bind('<FocusOut>', autocalcpprm)
        ppm.bind('<FocusOut>', autocalcpprm)
        win.mainloop()

def resetalldefaults():
        mscr.delete(0, END)
        sscr.delete(0, END)
        alcr.delete(0, END)
        cucr.delete(0, END)
        msr.delete(0, END)
        ssr.delete(0, END)
        qty.delete(0, END)
        
        mscr.insert(0, dfdv['Values'][0])
        sscr.insert(0, dfdv['Values'][1])
        alcr.insert(0, dfdv['Values'][2])
        cucr.insert(0, dfdv['Values'][3])
        msr.insert(0, dfdv['Values'][4])
        ssr.insert(0, dfdv['Values'][5])
        qty.insert(0, dfdv['Values'][6])

def settingsdef():
        def setting1def():
                try:
                        os.startfile(r'Default_Values.csv')
                except:
                        messagebox.showerror("Error", "Some files are missing!")

        def setting2def():
                try:
                        os.startfile(r'MM_CU_Feed_Rate.csv')
                except:
                        messagebox.showerror("Error", "Some files are missing!")

        def setting3def():
                try:
                        os.startfile(r'SS_AL_Feed_Rate.csv')
                except:
                        messagebox.showerror("Error", "Some files are missing!")
        
        wins = Toplevel()
        wins.geometry("900x900")
        
        setting1label = Label(wins, text = "1. Change Default Values (MSCR/SSCR/ALCR/CUCR/MSR/SSR/Qty)\n(Do not add/remove/change column/row name. Only change the values and save it directly.)", height = 4, font=300)
        setting1button = Button(wins, text = "Click Me to Edit DEFAULT VALUES", command = setting1def, cursor='hand2', bg='#e6e6e6')

        setting2label1 = Label(wins, text = "2. Change Feed Rate of MM and CU", height = 4, font=300)
        setting2label2 = Label(wins, text = "I)Do not add/remove/change column. II)You can add row for adding more flexibility of costing.\nIII)Costing should be in the asscending order of thickness. IV)Only change the values and save it directly.\nV)Any value of thickness is applicable for itself and lower than itself until its lower is defined.\nVI)The last value of thickness will apply for itselft and any value greater than that.", height = 4, font=300)
        setting2button = Button(wins, text = "Click Me to Edit FR - MM & CU", command = setting2def, cursor='hand2', bg='#e6e6e6')

        setting3label1 = Label(wins, text = "3. Change Feed Rate of SS and AL", height = 4, font=300)
        setting3label2 = Label(wins, text = "I)Do not add/remove/change column. II)You can add row for adding more flexibility of costing.\nIII)Costing should be in the asscending order of thickness. IV)Only change the values and save it directly.\nV)Any value of thickness is applicable for itself and lower than itself until its lower is defined.\nVI)The last value of thickness will apply for itselft and any value greater than that.", height = 4, font=300)
        setting3button = Button(wins, text = "Click Me to Edit FR - SS & AL", command = setting3def, cursor='hand2', bg='#e6e6e6')

        setting1label.pack()
        setting1button.pack()

        setting2label1.pack()
        setting2label2.pack()
        setting2button.pack()

        setting3label1.pack()
        setting3label2.pack()
        setting3button.pack()

        wins.mainloop()

#settingsdef()
# Create a File Explorer label
label_file_explorer = Label(root, text = "Sachin Industries", height = 4, fg = "blue", font=300)

frame4 = Frame(root)

frame4.columnconfigure(0, weight=1)
frame4.columnconfigure(1, weight=1)
frame4.columnconfigure(2, weight=1)
mcal = Button(frame4, text = 'Manual Calulator', command = mcalcu, bg='#e6e6e6')
browsef = Button(frame4,text = "Browse File",command = browseFiles, cursor='hand2', bg='#e6e6e6')
browsedir = Button(frame4,text = "Browse Folder",command = browseFolder, cursor='hand2', bg='#e6e6e6')

info = ''
text12 = Label(root, text=info)

textcname = Label(root, text="Enter Company Name")
cname = Entry(root, borderwidth=5, bg='white')

resetall = Button(root, text = 'Reset All To Default Values', command = resetalldefaults, bg='#ff928a')
settings = Button(root, command = settingsdef, image = settingimg)

frame2 = Frame(root)
frame2.columnconfigure(0, weight=1)
frame2.columnconfigure(1, weight=1)
frame2.columnconfigure(2, weight=1)
frame2.columnconfigure(3, weight=1)

textmscr = Label(frame2, text="MSCR")
mscr = Entry(frame2, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
textsscr = Label(frame2, text="SSCR")
sscr = Entry(frame2, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
textalcr = Label(frame2, text="Aluminium CR")
alcr = Entry(frame2, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
textcucr = Label(frame2, text="Copper CR")
cucr = Entry(frame2, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')

var1 = IntVar()
var2 = IntVar()
tick1 = Checkbutton(root, text="MS with Material", variable=var1, cursor='target')
tick2 = Checkbutton(root, text="SS with Material", variable=var2, cursor='target')

textmsr = Label(root, text="MSR")
msr = Entry(root, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')
textssr = Label(root, text="SSR")
ssr = Entry(root, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')

textqty = Label(root, text="Quantity")
qty = Entry(root, validate="key", validatecommand=(root.register(only_numbers), '%S'), borderwidth=5, bg='white')

go = Button(root, text = "Go",command = go, width = 20, bg = 'light green', cursor='hand2')




frame1 = Frame(root)

frame1.columnconfigure(0, weight=1)
frame1.columnconfigure(1, weight=1)
frame1.columnconfigure(2, weight=1)
frame1.columnconfigure(3, weight=1)
frame1.columnconfigure(4, weight=1)
frame1.columnconfigure(5, weight=1)
frame1.columnconfigure(6, weight=1)

v = StringVar(frame1, "1")

editfnametext = Label(frame1,text ='File Name')
editfnametext.grid(row = 0, column = 0)

editlmmaterialtext = Label(frame1,text ='Labour/With Material')
editlmmaterialtext.grid(row = 0, column = 1)

editmtext = Label(frame1,text ='Material')
editmtext.grid(row = 0, column = 2)

editthicktext = Label(frame1,text ='Thickness')
editthicktext.grid(row = 0, column = 3)

editqtytext = Label(frame1,text ='Quantity')
editqtytext.grid(row = 0, column = 4)

editppctext = Label(frame1,text ='Per Peice Cost')
editppctext.grid(row = 0, column = 5)

edittcosttext = Label(frame1,text ='Total Cost')
edittcosttext.grid(row = 0, column = 6)

tamttext = Label(frame1,text ='Total Amt')
tamttext.grid(row = 0, column = 7)
print('try = ', type(tamttext))


editfname = Entry(frame1, borderwidth=5, bg='white', font=10)
editfname.grid(row = 1, column = 0)
#Radiobutton(frame1, variable = v, value = '2').grid(row = 2, column = 0)

editlmmaterial = ttk.Combobox(frame1, width = 27, background='white', textvariable = tk.StringVar(), state= "readonly")
editlmmaterial['values']=('Labour Only','With Material','')
editlmmaterial.current(2)
editlmmaterial.grid(row = 1, column = 1)
#Radiobutton(frame1, variable = v, value = '3').grid(row = 2, column = 1)

editm = ttk.Combobox(frame1, width = 27, textvariable = tk.StringVar(), state= "readonly")
editm['values']=('MS','SS','CU','AL','')
editm.current(4)
editm.grid(row = 1, column = 2)
#Radiobutton(frame1, variable = v, value = '4').grid(row = 2, column = 2)

editthick = Entry(frame1, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
editthick.grid(row = 1, column = 3)
#Radiobutton(frame1, variable = v, value = '5').grid(row = 2, column = 3)

editqty = Entry(frame1, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_numbers), '%S'))
editqty.grid(row = 1, column = 4)
#Radiobutton(frame1, variable = v, value = '6').grid(row = 2, column = 4)

editppc = Entry(frame1, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
editppc.grid(row = 1, column = 5)
#Radiobutton(frame1, variable = v, value = '7').grid(row = 2, column = 5)

edittcost = Entry(frame1, borderwidth=5, bg='white', font=10, validate="key", validatecommand=(root.register(only_float), '%S', '%s', '%d', '%i'))
edittcost.grid(row = 1, column = 6)
#Radiobutton(frame1, variable = v, value = '8').grid(row = 2, column = 6)

tamt = Label(frame1,text ='0.0')
tamt.grid(row = 1, column = 7)
#Radiobutton(frame1, variable = v, value = '1').grid(row = 2, column = 7)

#functions
def addrecord():
        if editfname.get() == '' and editlmmaterial.get() == '' and editm.get() == '' and editthick.get() == '' and editqty.get() == '' and editppc.get() == '':
                messagebox.showwarning("Warning", "Enter some data to add record")
        else:
                if edittcost.get() == '':
                        tree.insert('', 'end', text="1", values=(editfname.get(), editlmmaterial.get(), editm.get(), editthick.get(), editqty.get(), editppc.get(), str(float(editqty.get())*float(editppc.get()))))
                        tamt.config(text = (float(tamt.cget('text'))+float(editppc.get())*float(editqty.get())))
                else:
                        tree.insert('', 'end', text="1", values=(editfname.get(), editlmmaterial.get(), editm.get(), editthick.get(), editqty.get(), editppc.get(), edittcost.get()))
                        tamt.config(text = (float(tamt.cget('text'))+float(edittcost.get())))
                #clear entry boxes
                editfname.delete(0, END)
                editlmmaterial.current('2')
                editm.current('4')
                editthick.delete(0, END)
                editqty.delete(0, END)
                editppc.delete(0, END)
                edittcost.delete(0, END)
                
def removeallrecord(e):
        print(e)
        for record in tree.get_children():
                tree.delete(record)
                tamt.config(text = '0.0')

def totalamtcalc():
        tamt.config(text = 0.0)
        for record in tree.get_children():
                values = tree.item(record, 'values')
                tamt.config(text = (float(tamt.cget('text'))+float(values[6])))

def removeselectedrecords(e):
        if len(tree.selection()) == 0:
                messagebox.showwarning("Warning", "Select a row to update")
        else:
                for record in tree.selection():
                        values = tree.item(record, 'values')
                        tree.delete(record)
                totalamtcalc()
'''
def selectrecord():
        if len(tree.selection()) == 0:
                messagebox.showwarning("Warning", "Select a row to update")
        elif len(tree.selection()) != 1:
                messagebox.showwarning("Warning", "Please Select only 1 row to be selected")
        else:
                #clear entry boxes
                editfname.delete(0, END)
                editlmmaterial.current('2')
                editm.current('4')
                editthick.delete(0, END)
                editqty.delete(0, END)
                editppc.delete(0, END)
                edittcost.delete(0, END)

                #grab record number
                selected = tree.focus()

                #grab record values
                values = tree.item(selected, 'values')
                
                #output to entry box
                editfname.insert(0, values[0])
                editthick.insert(0, values[3])
                editqty.insert(0, values[4])
                editppc.insert(0, values[5])
                edittcost.insert(0, values[6])

                if values[2] == 'MS':
                        editm.current(0)
                elif values[2] == 'SS':
                        editm.current(1)
                elif values[2] == 'AL':
                        editm.current(2)
                elif values[2] == 'CU':
                        editm.current(3)
'''             
def updaterecord2():
        global T
        global c
        global pp
        global m
        global sw
        
        global filename
        global verifyur
        global filepath
        c0,c1,c2,c3,c4,c5,c6,c7 = '','','','','','','',''
        for i in tree.selection():
                value = tree.item(i, 'values')
                if editfname.get() != '':
                        c0 = editfname.get()
                else:
                        c0 = value[0]
                
                if editlmmaterial.get() != '':
                        c1 = editlmmaterial.get()
                else:
                        c1 = value[1]
                
                if editm.get() != '':
                        c2 = editm.get()
                else:
                        c2 = value[2]
                
                if editthick.get() != '':
                        c3 = editthick.get()
                else:
                        c3 = value[3]
                
                if editqty.get() != '':
                        c4 = editqty.get()
                else:
                        c4 = value[4]
                
                if editppc.get() != '':
                        c5 = editppc.get()
                else:
                        c5 = value[5]
                
                if edittcost.get() != '':
                        c6 = edittcost.get()
                else:
                        c6 = value[6]

                c7 = value[7]
                
                if editlmmaterial.get() != '' or editm.get() != '' or editthick.get():
                        if value[7] == '':
                                messagebox.showerror("Error", "File Not Found!")
                                pass
                        else:
                                v1 = var1.get()
                                v2 = var2.get()
                                filepath = value[7]
                                filename = filepath[(filepath.rindex('/')+1):]
                                print('filepath=',filepath)
                                T = float(c3)
                                m = c2
                                data = open(filepath,"r+").read()
                                c = float(data[data.find("Cutting way: ")+13:][0:data[data.find("Cutting way: ")+13:].find(' ')])
                                pp = int(data[data.find("Pierce Qty: ")+12:][0:data[data.find("Pierce Qty: ")+12:].find(' ')])
                                sw = float(data[data.find("Sheet Weight: ")+14:][0:data[data.find("Sheet Weight: ")+14:].find(' ')])
                                tempq = qty.get()
                                qty.delete(0, END)
                                qty.insert(0, c4)
                                if c1 == 'Labour Only':
                                        tick1.deselect()
                                        tick2.deselect()
                                elif c1 == 'With Material':
                                        tick1.select()
                                        tick2.select()
                                else:
                                        print('Error in lmmaterial')
                                verifyur = 1
                                ls = algoauto()
                                print("New list: ",ls)
                                #to update record
                                tree.item(i, values = (c0, c1, ls[2], ls[3], ls[4], ls[5], ls[6], c7))
                                verifyur = 0
                                qty.delete(0, END)
                                qty.insert(0, tempq)
                                if v1 == 1:
                                        tick1.select()
                                else:
                                        tick1.deselect()
                                if v2 == 1:
                                        tick2.select()
                                else:
                                        tick2.deselect()
                elif editlmmaterial.get() == '' and editm.get() == '' and editthick.get() == '' and (editqty.get() != '' or editppc.get() != ''):
                        tree.item(i, values = (c0, c1, c2, c3, c4, c5, int(c4)*float(c5), c7))
                elif editlmmaterial.get() == '' and editm.get() == '' and editthick.get() == '' and (editqty.get() != '' or edittc.get() != ''):
                        tree.item(i, values = (c0, c1, c2, c3, c4, round(float(c6)/float(c4)), c6, c7))
                        
                                
                                
                

def updaterecord():
        if len(tree.selection()) == 0:
                messagebox.showwarning("Warning", "Select atleat 1 row to update")
        else:
                '''
                #grab record number
                selected = tree.focus()
                #save new data
                values = tree.item(selected, 'values')
                if len(tree.selection())==1:
                        if int(values [4])!=int(editqty.get()):
                                tree.item(selected, values = (editfname.get(), editlmmaterial.get(), editm.get(), editthick.get(), editqty.get(), editppc.get(), int(editqty.get())*float(editppc.get())))
                        else:
                                tree.item(selected, values = (editfname.get(), editlmmaterial.get(), editm.get(), editthick.get(), editqty.get(), editppc.get(), edittcost.get()))
                else:
                        for i in tree.selection():
                                value = tree.item(i, 'values')
                                if int(value[4])>=1:
                                        tree.item(i, values = (value[0], value[1], value[2], value[3], editqty.get(), editppc.get(), (float(value[5])/int(value[4]))*int(editqty.get())))
                                else:
                                        tree.item(i, values = (value[0], value[1], value[2], value[3], editqty.get(), editppc.get(), int(editqty.get())*float(value[5])))

                tamt.config(text = 0.0)
                for record in tree.get_children():
                        values = tree.item(record, 'values')
                        tamt.config(text = (float(tamt.cget('text'))+float(values[6])))
                                
                '''
                updaterecord2()
                #clear entry boxes
                editfname.delete(0, END)
                editlmmaterial.current('2')
                editm.current('4')
                editthick.delete(0, END)
                editqty.delete(0, END)
                editppc.delete(0, END)
                edittcost.delete(0, END)

                totalamtcalc()
                
def clearentryboxesdef():
        editfname.delete(0, END)
        editlmmaterial.current('2')
        editm.current('4')
        editthick.delete(0, END)
        editqty.delete(0, END)
        editppc.delete(0, END)
        edittcost.delete(0, END)

'''def trialtest123():
        print('var1=', var1)
        print('var2=', var2)
        v1 = var1
        v2 = var2
        
        if v1 == 1:
                tick1.deselect()
        else:
                tick1.select()
        if v2 == 1:
                tick2.deselect()
        else:
                tick2.select()'''

frame3 = Frame(root)

frame3.columnconfigure(0, weight=1)
frame3.columnconfigure(1, weight=1)
frame3.columnconfigure(2, weight=1)
frame3.columnconfigure(3, weight=1)
frame3.columnconfigure(4, weight=1)
frame3.columnconfigure(5, weight=1)
#frame3.columnconfigure(6, weight=1)





addrec = Button(frame3, text = 'Add Record', command = addrecord, bg='#e6e6e6')
removeallrec = Button(frame3, text = 'Remove All Records', command = lambda:removeallrecord, bg='#e6e6e6')
removeselectedrecs = Button(frame3, text = 'Remove Selected Record(s)', command = lambda:removeselectedrecords, bg='#e6e6e6')
#selectrec = Button(frame3, text = 'Select Record', command = selectrecord, bg='#e6e6e6')
clearentryboxes = Button(frame3, text = 'Clear all Entry boxes', command = clearentryboxesdef, bg='#e6e6e6')
updaterec = Button(frame3, text = 'Update Record', command = updaterecord, bg='#e6e6e6')
export = Button(frame3, text = 'Export To Excel', command = exportxl, bg='#e6e6e6')
'''test = Button(frame3, text = 'test', command = trialtest123, bg='#e6e6e6')
test.grid(column = 7, row = 0, pady=10)'''

def maindraganddrop(event):
    global filepath
    global filename
    global last
    global folderpath
    print(event.data.count('{'),'no. of files')
    mstr = event.data
    for i in range(0,event.data.count('{')):
        xstr = mstr[mstr.find('{')+1:mstr.find('}')]
        mstr = mstr[mstr.find('}')+1:]
        if xstr.endswith('.Hnf'):
            filepath = xstr
            last = 'file'
            if filepath != '':
                    folderpath = filepath[:filepath.rindex('/')]
                    filename = filepath[(filepath.rindex('/')+1):]
                    # Change label contents
                    label_file_explorer.configure(text="File Opened: "+filepath)
                    text12.config(text = "Reading File: " + filename)
            valuesfromfile()
    cnamechange()
    print('yehhhhhh')
    print(event)
    print(event.data)

# store the filename associated with each tree item in a dictionary
tree.filenames = {}
# add a boolean flag to the tree which can be used to disable files from the tree being dropped on the tree again
tree.dragging = False

tree.drop_target_register(DND_FILES)
tree.dnd_bind('<<Drop>>', maindraganddrop)

# drag methods

def drag_init(event):
    data = ()
    sel = tree.select_item()
    if sel:
        # in a decent application we should check here if the mouse
        # actually hit an item, but for now we will stick with this
        data = (tree.filenames[sel],)
        tree.dragging = True
        return ((ASK, COPY), (DND_FILES, DND_TEXT), data)
    else:
        # don't start a dnd-operation when nothing is selected; the
        # return "break" here is only cosmetical, return "foobar" would
        # probably do the same
        return 'break'

def drag_end(event):
    # reset the "dragging" flag to enable drops again
    tree.dragging = False

tree.drag_source_register(1, DND_FILES)
tree.dnd_bind('<<DragInitCmd>>', drag_init)
tree.dnd_bind('<<DragEndCmd>>', drag_end)


# Grid method is chosen for placing the widgets at respective positions in a table like structure by specifying rows and columns

settings.grid(column = 0, row = 2, sticky='w')

frame4.grid(column = 0, row = 2, columnspan=5)
mcal.grid(column = 2, row = 0)
browsef.grid(column = 0, row = 0)
browsedir.grid(column = 1, row = 0)

text12.grid(column = 0, row = 3, columnspan=5)

resetall.grid(column = 0, row = 4, columnspan=5)

frame2.grid(column=0, row=5, columnspan=15, sticky="nsew")
textmscr.grid(column = 0, row = 0)
textsscr.grid(column = 1, row = 0)
textalcr.grid(column = 2, row = 0)
textcucr.grid(column = 3, row = 0)
mscr.grid(column = 0, row = 1)
sscr.grid(column = 1, row = 1)
alcr.grid(column = 2, row = 1)
cucr.grid(column = 3, row = 1)

mscr.insert(0, dfdv['Values'][0])
sscr.insert(0, dfdv['Values'][1])
alcr.insert(0, dfdv['Values'][2])
cucr.insert(0, dfdv['Values'][3])

tick1.grid(column = 0, row = 8, columnspan = 5)
tick2.grid(column = 0, row = 9, columnspan = 5)

textmsr.grid(column = 0, row = 10, sticky="e")
msr.grid(column = 1, row = 10, sticky="w")
msr.insert(0, dfdv['Values'][4])
textssr.grid(column = 0, row = 11, sticky="e")
ssr.grid(column = 1, row = 11, sticky="w")
ssr.insert(0, dfdv['Values'][5])

textqty.grid(column = 0, row = 12, sticky="e")
qty.grid(column = 1, row = 12, sticky="w")
qty.insert(0, dfdv['Values'][6])

textcname.grid(column = 0, row = 13, sticky="e")
cname.grid(column = 1, row = 13, sticky="w")

go.grid(column = 0,row = 14, columnspan=5)

#table position
tree.grid(column=0, row=15, sticky='nsew', columnspan=5)
scrollbar.grid(column=5, row=15, sticky='nsw')

frame1.grid(column=0, row=16, columnspan=15, sticky="nsew")
frame3.grid(column=0, row=17, columnspan=15, sticky="nsew")

addrec.grid(column = 0, row = 0, pady=10)
removeallrec.grid(column = 1, row = 0, pady=10)
removeselectedrecs.grid(column = 2, row = 0, pady=10)
#selectrec.grid(column = 3, row = 0, pady=10)
updaterec.grid(column = 3, row = 0, pady=10)
clearentryboxes.grid(column = 4, row = 0, pady=10)
export.grid(column = 5, row = 0, pady=10)

#tree.drop_target_register(DND_FILES)
#tree.dnd_bind('<<DropEnter>>', testdrop)

def newtest(char):
        print('It worked!!!'*30)

tree.bind('<Return>', newtest)
removeselectedrecs.bind('<Double-1>', removeselectedrecords)
removeallrec.bind('<Double-1>', removeallrecord)
# Let the root wait for any events
root.mainloop()
