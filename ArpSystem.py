# !/usr/bin/env python
# _*_ encoding:utf-8 _*_

from Tkinter import *
import tkMessageBox
import os
import time
import xlwt
import xlrd
from xlutils.copy import copy

__BaseDir__ = os.path.dirname(os.path.abspath(__file__))
__SystemInfo__ = ' CopyRight：Mr.Liu  SystemVersion：1.0  QQ：345919932'

MODELS = {
    '工作号': '',
    '线材类型': '',
    '高压线圈数': '',
    '低压线圈数': '',
    '高压线圈型号': '',
    '低压线圈型号': '',
    '铁芯数量': '',
    '铁芯编号': '',
    '电频总装': '',
    '装配日期': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())),
    '标号': '',
    '装配成功数量': 0,
    '高压线圈使用数': 0,
    '低压线圈使用数': 0,
    '铁芯使用数': 0,
}

HIGH = 3
LOW = 3
IROM = 1

def CreateExcel():
    FileName = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    BaseFile = os.path.join(__BaseDir__, '%s.xls' % FileName)
    workbook = ''
    table = ''
    if os.path.exists(BaseFile):
        workbook = xlrd.open_workbook(BaseFile)
        workbook.encoding = 'utf-8'
        table = workbook.sheet_by_index(0)
        changeBook = copy(workbook)
        sheet = changeBook.get_sheet(0)
        row = 0
        col = 0
        if table.nrows == 0:
            for top in MODELS.keys():
                sheet.write(row, col, top.decode('utf-8'))
                col +=1
        col = 0
        for data in MODELS.values():
            sheet.write(table.nrows, col, data)
            col +=1
        changeBook.save(BaseFile)
    else:
        workbook = xlwt.Workbook(encoding='utf-8')
        table = workbook.add_sheet('Resource', cell_overwrite_ok=True)
        workbook.save(BaseFile)

def ShowInfo():
    if heightWireNum.get() == '' or heightWire.get() == 0:
        tkMessageBox.showinfo('值不能为空', '高压线圈数量不能为空!')
        text.insert(END, '装配失败...\r\n')

    if lowWireNum.get() == '' or lowWire.get() == 0:
        tkMessageBox.showinfo('值不能为空', '低压线圈数量不能为空!')
        text.insert(END, '装配失败...\r\n')

    if ironcoreNum.get() == '' or lowWire.get() == 0:
        tkMessageBox.showinfo('值不能为空', '铁芯数量不能为空!')
        text.insert(END, '装配失败...\r\n')
    MODELS['工作号'] = jobNumber.get()
    MODELS['线材类型'] = wireType.get()
    MODELS['高压线圈数'] = heightWire.get()
    MODELS['低压线圈数'] = lowWire.get()
    MODELS['高压线圈型号'] = heightWireNumber.get()
    MODELS['低压线圈型号'] = lowWireNumber.get()
    MODELS['电频总装'] = Assembly.get()
    MODELS['标号'] = markNumber.get()
    MODELS['铁芯数量'] = ironCore.get()
    MODELS['铁芯编号'] = ironcoreModel.get()

def BeginAssembly():
    global HIGHCOUNT,LOWCOUNT,IROMCOUNT,COUNT
    text.delete(1.0, END)
    text.insert(END, '开始装配...\r\n')
    text.insert(END, '正在装配请稍后...\r\n')
    ShowInfo()
    if MODELS['高压线圈数'] < 3 or MODELS['低压线圈数'] < 3 or MODELS['铁芯数量'] < 1:
        tkMessageBox.showinfo('材料不足', '材料不足')
        text.insert(END, '装配失败...\r\n')
    else:
        text.insert(END, '正在配对材料请稍后...\r\n')
        MODELS['高压线圈数'] -= HIGH
        MODELS['低压线圈数'] -= LOW
        MODELS['铁芯数量'] -= IROM
        MODELS['高压线圈使用数'] += HIGH
        MODELS['低压线圈使用数'] += LOW
        MODELS['铁芯使用数'] += IROM
        MODELS['装配成功数量'] +=1
        heightWire.set(MODELS['高压线圈数'])
        lowWire.set(MODELS['低压线圈数'])
        ironCore.set(MODELS['铁芯数量'])
        text.insert(END, '成功装配数量: %s\r\n' % MODELS['装配成功数量'])
        text.insert(END, '装配成功...')
        CreateExcel()
# def main():
root = Tk()
root.title('ErpSystem')
root.geometry('600x300+100+100')
root.resizable(width=True, height=True)

frame = Frame(root)
frame.pack(padx=10,pady=10)
Label(frame, text='工作号:', pady=5).grid(row=0, column=0)
jobNumber = Entry(frame,width=15)
jobNumber.grid(row=0,column=1)

Label(frame, text='电频总装:', pady=5).grid(row=0, column=2)
Assembly = Entry(frame,width=15)
Assembly.grid(row=0,column=3)

Label(frame, text='线圈类型:', padx=5, pady=5).grid(row=0,column=4)
wireType = Entry(frame,width=15)
wireType.grid(row=0,column=5)

Label(frame, text='高压线圈:',padx=5, pady=5).grid(row=1,column=0)
heightWire = IntVar()
heightWire.set('')
heightWireNum = Entry(frame,width=15, textvariable=heightWire)
heightWireNum.grid(row=1,column=1)

Label(frame, text='低压线圈:',padx=5, pady=5).grid(row=1,column=2)
lowWire = IntVar()
lowWire.set('')
lowWireNum = Entry(frame,width=15,textvariable=lowWire)
lowWireNum.grid(row=1,column=3)

Label(frame, text='高压线圈编号:', padx=5, pady=5).grid(row=1,column=4)
heightWireNumber = Entry(frame,width=15)
heightWireNumber.grid(row=1,column=5)

# Label(frame, text='线圈电频:',padx=5).grid(row=1,column=4)
# Electricfrequency = [
#     ('高', 1),
#     ('低', 0)
# ]
# v = StringVar()
# v.set(1)
# n = 5
# for text, mode in Electricfrequency:
#     radio = Radiobutton(frame,text=text, variable=v, value=mode, padx=5)
#     radio.grid(row=1, column=n)
#     n +=1
Label(frame, text='低压线圈编号:', padx=5, pady=5).grid(row=2,column=0)
lowWireNumber = Entry(frame,width=15)
lowWireNumber.grid(row=2,column=1)

Label(frame, text='铁芯数:',padx=5, pady=5).grid(row=2, column=2)
ironCore = IntVar()
ironCore.set('')
ironcoreNum = Entry(frame,width=15,textvariable=ironCore)
ironcoreNum.grid(row=2,column=3)

Label(frame, text='铁芯编号(1-3):',padx=5, pady=5).grid(row=2, column=4)
ironcoreModel = Entry(frame,width=15)
ironcoreModel.grid(row=2,column=5)

Label(frame, text='标号:',padx=5, pady=5).grid(row=3)
markNumber = Entry(frame,width=15)
markNumber.grid(row=3, column=1)

start = Button(root, text='开始装配', command=BeginAssembly, relief=RAISED)
start.pack(pady=5)

text = Text(root, height=20 ,undo=True, borderwidth=2, relief=RIDGE)
text.pack(expand=YES, fill=BOTH)

scroll = Scrollbar(text)
text.config(yscrollcommand=scroll.set)
scroll.config(command=text.yview)
scroll.pack(side=RIGHT, fill=Y)

status = Label(root,text='%s' % __SystemInfo__, bd=1, relief=SUNKEN, anchor=W)
status.pack(side=BOTTOM, fill=X)
root.mainloop()

# if __name__ == "__main__":
#     main()

