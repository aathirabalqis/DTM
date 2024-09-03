import os
import math
import pandas as pd 
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.scrolledtext import ScrolledText
from tkcalendar import Calendar, DateEntry

root = Tk()

root.title('DTM Analysis')
root.geometry('425x580')

def getid():
    global all_ids, ind, out, filepath
    
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    p1.config(text = 'Path: ' + str(filenames[0]))
    
    filepath = '\\'.join(filenames[0].split('/')[:-1])
    out = filepath + '\\' + 'Seal Data - Compiled.xlsx'
    df = pd.DataFrame()
    print('start')
    
    # writer = pd.ExcelWriter(filenames[0], engine = 'xlsxwriter')
    file = pd.ExcelFile(filenames[0])
    sheets = file.sheet_names
    
    # for i in sheets:
        # temdf = pd.read_excel(filenames[0], sheet_name = i)
        # temdf['PE'] = i
        # df = pd.concat([df,temdf]) if len(df) > 5 else temdf
    
    # writer.close()
    df = pd.concat(pd.read_excel(filenames[0], sheet_name=None), ignore_index=True)
    if os.path.exists(out) == False: df.to_excel(out, index = False)
    
    # output meter IDs with no installation
    noint = df.loc[(pd.isnull(df['Installation ID'])) & (df['Meter No.'].str.len() > 10), 'Meter No.'].drop_duplicates().dropna().astype(str).tolist()
                   
    for i in noint:
        textbox.insert(END, i + '\n')
    
    total.config(text = 'Displayed IDs = ' + str(len(noint)))
    outp.config(text = 'Paste into Serial Number to find missing Inst. ID')

def getinst():
    global out
    
    textbox.delete('1.0', END)
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    p2.config(text = 'Path: ' + str(filenames[0]))
    
    
    inst = pd.read_excel(filenames[0])
    
    writer = pd.ExcelWriter(out,  mode = 'a', engine = 'openpyxl')
    
    outdf = pd.read_excel(out, sheet_name = 'Compiled') if 'Compiled' in writer.book.sheetnames else pd.read_excel(out) 
   
    if 'Contract Account' not in outdf.columns:
        print('first')
        inst = inst.rename(columns = {'Serial Number':'Meter No.'})
        outdf = pd.merge(left = outdf, right = inst, how = 'left', on = 'Meter No.')
        outdf.loc[(pd.isnull(outdf['Installation ID'])), 'Installation ID'] = outdf['Installation']
        outdf = outdf.drop(columns = ['Installation'])
        print(outdf)
        
        
        # output inst
        output = outdf['Installation ID'].drop_duplicates().dropna().astype(str).tolist()
        outp.config(text = 'Paste into Installation to find missing CA')
        p2.config(text = 'Path: ')

        
    else:
        print('second')

        inst = inst.rename(columns = {'Installation':'Installation ID','Contract Account':'CA'})
        outdf = pd.merge(left = outdf, right = inst, how = 'left', on = 'Installation ID')
        outdf.loc[(pd.isnull(outdf['Contract Account'])), 'Contract Account'] = outdf['CA']
        outdf = outdf.drop(columns = ['CA'])
        
        #fill in new meter number col with old values
        outdf.loc[(pd.isnull(outdf['Meter No.'])), 'Serial Number'] = outdf['Meter No.']
        
        output = outdf['Contract Account'].drop_duplicates().dropna().astype(str).tolist()
        outp.config(text = 'Get kWh readings with SQL & BCRM ->')
        outdf = outdf[['Substation name','Installation ID','Meter No.','Serial Number','Contract Account']]
        
        clout = outdf.loc[(pd.isnull(outdf['Contract Account']) == False) & (pd.isnull(outdf['Serial Number']) == False)]
        clout = clout.drop(columns = ['Meter No.'])
        clout.to_excel(writer, sheet_name = 'Clean', index = False)
        writer.book.remove(writer.book['Compiled'])
        

        textvar2.set('BCRM')
        
    for i in output:
        if '.' in i: i = i[:-2]
        textbox.insert(END, i + '\n')

    total.config(text = 'Displayed IDs = ' + str(len(output)))
    outdf.to_excel(writer, sheet_name = 'Compiled', index = False)
    writer.close()
   
def getcons():
    global out, all_ids, ind
    
    textbox.delete('1.0', END)
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    p3.config(text = 'Path: ' + str(filenames[0]))
    
    zbi = pd.read_excel(filenames[0])
    
    noofdays = int((cal2.get_date() - cal.get_date()).days)
    
    zbi['BCRM Reading'] = (zbi['Current usage consumption'].astype(float) / zbi['Bill duration'].astype(int) * noofdays).astype(int)
    zbi = zbi.loc[(zbi['Print Date'].dt.month == cal.get_date().month) & (zbi['Print Date'].dt.year == cal.get_date().year), ['Contract Account','BCRM Reading']]
    print(zbi)
    
    writer = pd.ExcelWriter(out,  mode = 'a', engine = 'openpyxl')
    outdf = pd.read_excel(out, sheet_name = 'Clean')
    
    outdf = pd.merge(left = outdf, right = zbi, how = 'left', on = 'Contract Account')
    
    writer.book.remove(writer.book['Clean'])
    outdf.to_excel(writer, sheet_name = 'Clean', index = False)
    
    writer.close()
    all_ids = []
    ind = 0
    
    # output meter IDs for SQL
    if 'SQL Diff.' in outdf.columns: textbox.insert(END,'Click button to generate report!')  
    
    else:
        outdf = outdf.loc[(outdf['Serial Number'].str.len() > 15)]
        new_ids = outdf['Serial Number'].drop_duplicates().astype(str).tolist()
        all_ids.append(new_ids)
        act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
        
        if len(new_ids) > 1000: divide(new_ids, 1000)
        textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')  
        total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
        outp.config(text = 'Find kWh reading on SQL & upload for both dates')
        textvar2.set('SQL')
    
def getsql():
    
    global out
    
    textbox.delete('1.0', END)
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    p4.config(text = 'Path: ' + str(filenames))
    
    writer = pd.ExcelWriter(out,  mode = 'a', engine = 'openpyxl')
    outdf = pd.read_excel(out, sheet_name = 'Clean')
    bigsq = pd.DataFrame()

    for i in filenames:
        sq = pd.read_excel(i)
        print(len(sq))
        
        # reformat
        date = pd.to_datetime(sq.loc[1,'READ_TIME'].split(' ')[0], format = '%d-%b-%y').strftime('%d/%m/%Y')
        sq = sq[['METER_ID','READ_VALUE']]
        sq = sq.rename(columns = {'METER_ID':'Serial Number','READ_VALUE':date})
        # biqsq = pd.concat([bigsq,sq]) if len(bigsq) > 5 else sq
        outdf = pd.merge(left = outdf, right = sq, how = 'left', on = 'Serial Number')
        
        for j in outdf.columns:
            if '_x' in j: 
                outdf.loc[(pd.isnull(outdf[i])), i] = outdf[i.split('_')[0] + '_y']
                outdf = outdf.rename(columns = {i:i.split('_')[0]})
                outdf = outdf.drop(columns = [i.split('_')[0] + '_y'])
        
    # print(bigsq)
    # outdf = pd.merge(left = outdf, right = bigsq, how = 'left', on = 'Serial Number')
    if cal.get_date().strftime('%d/%m/%Y') in outdf.columns and cal2.get_date().strftime('%d/%m/%Y') in outdf.columns:
        
        # is it okay if null values are put as 0? should be ??
        
    
        outdf['SQL Diff.'] = (outdf[cal2.get_date().strftime('%d/%m/%Y')].astype(int) - outdf[cal.get_date().strftime('%d/%m/%Y')])
        outdf.loc[(pd.isnull(outdf['SQL Diff.']) == False), 'SQL Reading'] = outdf['SQL Reading'].astype(int)
        
    if 'BCRM Reading' in outdf.columns: textbox.insert(END,'Click button to generate report!') 
            
    else: 
        output = outdf['Contract Account'].drop_duplicates().dropna().astype(str).tolist()
        for i in output:
            if '.' in i: i = i[:-2]
            textbox.insert(END, i + '\n')
            
        writer.book.remove(writer.book['Clean'])
        outdf.to_excel(writer, sheet_name = 'Clean', index = False)

    writer.close()    
    
def genrep():
    global out, filepath
    
    # read seal data compiled
    df = pd.read_excel(out, sheet_name = 'Clean')
    finalout = filepath + '\\' + 'DTM Analysis.xlsx'
    
    # iterate through each PE and make a sheet for each with columns installation, meter no, kwh reading - later add on dtm reading\\
    for i in df['Substation name'].drop_duplicates().tolist():
        pedf = df.loc[(df['Substation name'] == i)]
        pedf = pedf.rename(columns = {'SQL Reading':'kWh Reading'})
        pedf.loc[(pd.isnull(pedf['kWh Reading'])), 'kWh Reading'] = pedf['BCRM Reading']
        
        if os.path.exists(finalout): 
            writer = pd.ExcelWriter(finalout, mode = 'a', engine = 'openpyxl')
            pedf.to_excel(writer, sheet_name = i, index = False)
            writer.close()
        else: pedf.to_excel(finalout, sheet_name = i, index = False)
            
    
    # make a summary sheet of each PEs, sum of kwh, sum of dtm, diff and column
    
    print('test')

def divide(array,lim):

    global all_ids, ind

    all_ids = []
    for i in range(math.ceil(len(array)/lim)):
        temp = array[lim*i:lim*(i+1)]
        all_ids.append(temp)
    
    b4["state"] = NORMAL   
    return all_ids
    
def nextt():

    global all_ids, ind
       
    b3["state"] = NORMAL
    b4["state"] = DISABLED
    ind += 1
    
    textbox.delete('1.0', END)
    
    if textvar2.get() == 'BCRM':
        for i in all_ids[ind]:
            if '.' in i: i = i[:-2]
            textbox.insert(END, i + '\n')
        outp.config(text = outpp[ind])
            
    else: textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')  
    
    if ind < len(all_ids)-1:
        b4["state"] = NORMAL

    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
   
def back():
    global all_ids, ind
    
    b4["state"] = NORMAL
    b3["state"] = DISABLED
    ind -= 1
    
    textbox.delete('1.0', END)
    
    if textvar2.get() == 'BCRM':
        for i in all_ids[ind]:
            if '.' in i: i = i[:-2]
            textbox.insert(END, i + '\n')
        outp.config(text = outpp[ind])
        
    else: textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')
    
    if ind != 0:
       b3["state"] = NORMAL 
   
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))

def outid2():

    global out, all_ids, ind

    outdf = pd.read_excel(out, sheet_name = 'Clean')
    ind = 0
    all_ids = []
    textbox.delete('1.0', END)
    
    if textvar2.get() == 'SQL':
        outdf = outdf.loc[(outdf['Serial Number'].str.len() > 15)]
        new_ids = outdf['Serial Number'].drop_duplicates().astype(str).tolist()
        all_ids.append(new_ids)
        act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
        textvar2.set('SQL')
        
        
        if len(new_ids) > 1000: divide(new_ids, 1000)
            
        
        textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')  
        
        total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
        outp.config(text = 'Find kWh reading on SQL & upload for both dates')
        
    else:
        new_ids = outdf['Contract Account'].drop_duplicates().dropna().astype(str).tolist() 
       
        for i in new_ids:
            if '.' in i: i = i[:-2]
            textbox.insert(END, i + '\n')
        
        total.config(text = 'Displayed IDs = ' + str(len(new_ids)))
        outp.config(text = 'Find kWh reading on BCRM ZBI and upload')

def comb(event):
    outid2()

def exitt(event):
    root.quit()
	
def test():
    global out
    print('start')
    print(int((cal2.get_date() - cal.get_date()).days))
    
    df = pd.read_excel(out, sheet_name = 'Clean')
    print(df.columns)    


Button(root, text = 'Get Meter IDs', command = getid).place(x = 35, y = 70) 
Button(root, text = 'Get Missing Data', command = getinst).place(x = 23, y = 110) 
Button(root, text = 'Get Cons. (kWh)', command = getcons).place(x = 28, y = 150) 
Button(root, text = 'Get SQL Reading', command = getsql).place(x = 26, y = 190) 
Button(root, text = 'Generate DTM Analysis Report ', command = genrep).place(x = 130, y = 235) #report

p1 = Label(root, text = 'Path: ')
p2 = Label(root, text = 'Path: ')
p3 = Label(root, text = 'Path: ')
p4 = Label(root, text = 'Path: ')

Label(root, text = 'Start Date: ').place(x = 75, y = 30)
Label(root, text = 'End Date: ').place(x = 210, y = 30)

# cal = Calendar(root)
cal = DateEntry(root, width = 7)
cal.place(x = 135, y = 30)

cal2 = DateEntry(root, width = 7)
cal2.place(x = 270, y = 30)

b3 = Button(root, text = 'Back', command = back)
b4 = Button(root, text = 'Next', command = nextt)

p1.place(x = 130, y  = 72)
p2.place(x = 130, y  = 112)
p3.place(x = 130, y  = 152)
p4.place(x = 130, y  = 192)
b3.place(x = 30, y = 537)
b4.place(x = 67, y = 537)

textvar2 = StringVar()
choices = ['SQL','BCRM']
outpp = ['Paste into Serial Number to find CA','Paste into Installation to find CA']

cb = ttk.Combobox(root, textvariable = textvar2, values = choices, width = 6)
cb.bind('<<ComboboxSelected>>', comb)
cb.place(x = 315, y = 280)

outp = Label(root, text = 'Output')
total = Label(root, text = '')
act_total = Label(root, text = '')

outp.place(x = 30, y = 280)
total.place(x = 265, y = 540)
act_total.place(x = 105, y = 540)

b3["state"] = DISABLED
b4["state"] = DISABLED

textbox = ScrolledText(root, height = 14, width = 43)
textbox.place(x = 30, y = 306)

root.resizable(True, False) 
root.bind('<Escape>', exitt)
root.mainloop()
