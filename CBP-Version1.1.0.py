# -*- coding: utf-8 -*-
from Tkinter import *
import ttk
import tkMessageBox
import win32com.client
import os
import sys
import decimal
# callback Functions
def callback1():
    # 调用filedialog模块的askdirectory()函数去打开文件夹,调用filedialog模块的askopenfilename()函数去打开文件
    # 目标文件夹路径
    global folderpath
    folderpath = tkFileDialog.askdirectory()
    entry1.delete(0, END)  # 清空entry里面的内容
    # 2016/5/5修改以下
    if folderpath:
        entry1.insert(0, folderpath)  # 将选择好的路径加入到entry里面
        # 读取目标路径下的结果表,然后以checkbutton的形式动态第显示每一个sheet,checkbutton默认为勾选状态
        global TTwb
        winwbapp = win32com.client.Dispatch('Excel.Application')
        winwbapp.Visible = False
        TTwb = winwbapp.Workbooks.Open(folderpath + '/result.xlsx')
        count = TTwb.Sheets.Count
        # 2016/5/9结果表的sheet数量修改怎么一部判断，result表中有其他的sheet，判断是不是componentsheet
        # 遍历结果表中的sheet，识别出哪些是componentsheet,新建一个listL存放结果表中所有sheet的名字
        global L
        L = []
        for selects in range(1,count+1):
            TTwbs = TTwb.Sheets(selects)
            if TTwbs.Cells(1, 1).Value == 'Program':
                L.append(TTwbs.Name)
            else:
                pass
        print L
        TTwb.Close()
        winwbapp.Quit()
        # 初始化checklist,默认是勾选的,首先初始化一个mapcb
        global mapcb
        mapcb = {}
        for mapkey in L:
            mapcb[mapkey] = 1
        i = -1
        for key in mapcb:
            mapcb[key] = IntVar()
            global checkb
            checkb = Checkbutton(root, text=key, variable=mapcb[key])
            mapcb[key].set(1)
            i = i + 1
            checkb.grid(row=2 + i, column=1, sticky=W)
    else:
        # 弹出警告please select path
        tkMessageBox.showwarning("Warning", "Please select FilePath first!")
        # 2016/5/5修改以上
def callback2():
    import win32com.client
    # step1
    # 在点击Creat Button之前，增加一步判断,entry1中的值不能为空
    if e1.get()=='':
        tkMessageBox.showwarning("Warning", "Please select FilePath first befort click creat button!")
    # 2016/5/5修改以下
    else:
        # step2
        # 创建一个列表类型list Z[]用来存放，被选中的sheet的名字
        Z = []
        for key, value in mapcb.items():
            state = value.get()
            if state != 0:
                Z.append(key)
        # 4/19/2016判断选择的sheet是否W为空
        if Z == []:
            tkMessageBox.showerror("Error", "Please Select at least one sheet!")
        else:
            #2016/5/12增加一步对Rules的判断
            # 2016/5/12增加一个全局变量logt存放error log,logr代表写入的行数，即出现错误的个数
            global logr
            logr = 0
            global logt
            logt = open(folderpath + '/log.txt', 'w')
            
            # 2016/5/12增加一个全局变量logt存放error log
            # winwbapp = win32com.client.Dispatch('Excel.Application')
            # winwbapp.Visible = False
            # TRule12 = winwbapp.Workbooks.Open(folderpath + '/result.xlsx')
            # TRtemp = winwbapp.Workbooks.Open(folderpath + '/Rules.xlsx')
            # for Li in Z:
            #     if Li == 'SD&Others_Daily Resource Track':
            #         pass
            #     elif Li == 'Mobile_Daily Resource Tracking':
            #         # Rule1&Rule4
            #         # Rule1
            #         TRule12s = TRule12.Sheets(Li)
            #         Rows1 = TRule12s.UsedRange.Rows.Count
            #         TRtemps = TRtemp.Sheets('Rule1')
            #         Rows2 = TRtemps.UsedRange.Rows.Count
            #         NewRows1 = Rows1
            #         for pd in range(1, Rows1 + 1):
            #             if TRule12s.Cells(pd, 1).Value == None:
            #                 NewRows1 = pd - 1
            #                 break
            #             else:
            #                 pass
            #         if NewRows1 == 1:
            #             pass
            #         else:
            #             value = []
            #             for rows2 in range(2, Rows2 + 1):
            #                 if TRtemps.Cells(rows2, 1).Value == Li:
            #                     value.append(TRtemps.Cells(rows2, 2).Value)
            #                 else:
            #                     pass
            #             for i in range(2, NewRows1 + 1):
            #                 if TRule12s.Cells(i, 11).Value in value:
            #                     pass
            #                 else:
            #                     # 不符合rule,数据字体标红
            #                     # TRule12s.Cells(i, 11).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 11).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(11)
            #                     logt.write(log + '\n')
            #             # Rule4实现
            #             TRtemps2 = TRtemp.Sheets('Rule4')
            #             # value1-TestAsset-2,value2-Responsible Manager-7,value3-Purchase Order-8
            #             # value4-Target Milestone-4,value5-Target Launch-5,value6-Firmware Revision-6
            #             for i in range(2, NewRows1 + 1):
            #                 value1 = TRtemps2.Cells(2, 1).Value
            #                 value2 = TRtemps2.Cells(2, 2).Value
            #                 value3 = TRtemps2.Cells(2, 3).Value
            #                 if TRule12s.Cells(i, 2).Value == value1:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 2).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 2).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(2) + ' ' + 'correct value should be ' + value1
            #                     logt.write(log + '\n')
            #                 if TRule12s.Cells(i, 7).Value == value2:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 7).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 7).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(7) + ' ' + 'correct value should be ' + value2
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 8).Value == value3:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 8).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 8).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(8) + ' ' + 'correct value should be ' + value3
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 4).Value == None or TRule12s.Cells(i, 4).Value == 'N/A':
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 4).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 4).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(
            #                         4) + ' ' + 'correct value should be ' + 'None or N/A'
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 5).Value == None or TRule12s.Cells(i, 5).Value == 'N/A':
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 5).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 5).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(
            #                         5) + ' ' + 'correct value should be ' + 'None or N/A'
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 6).Value == None or TRule12s.Cells(i, 6).Value == 'N/A':
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 6).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 6).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(
            #                         6) + ' ' + 'correct value should be ' + 'None or N/A'
            #                     logt.write(log + '\n')
            #                     # Rule4实现
            #     elif Li == 'Ink Sub_Daily Resource Tracking':
            #         # Rule1&Rule5
            #         # Rule1
            #         TRule12s = TRule12.Sheets(Li)
            #         Rows1 = TRule12s.UsedRange.Rows.Count
            #         TRtemps = TRtemp.Sheets('Rule1')
            #         Rows2 = TRtemps.UsedRange.Rows.Count
            #         NewRows1 = Rows1
            #         for pd in range(1, Rows1 + 1):
            #             if TRule12s.Cells(pd, 1).Value == None:
            #                 NewRows1 = pd - 1
            #                 break
            #             else:
            #                 pass
            #         if NewRows1 == 1:
            #             pass
            #         else:
            #             value = []
            #             for rows2 in range(2, Rows2 + 1):
            #                 if TRtemps.Cells(rows2, 1).Value == Li:
            #                     value.append(TRtemps.Cells(rows2, 2).Value)
            #                 else:
            #                     pass
            #             for i in range(2, NewRows1 + 1):
            #                 if TRule12s.Cells(i, 11).Value in value:
            #                     pass
            #                 else:
            #                     # 不符合rule,数据字体标红
            #                     # TRule12s.Cells(i, 11).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 11).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(11)
            #                     logt.write(log + '\n')
            #             # 实现Rule5
            #             TRtemps2 = TRtemp.Sheets('Rule5')
            #             # value1-TestAsset-2,value2-Responsible Manager-7,value3-Purchase Order-8
            #             # value4-Target Milestone-4,value5-Target Launch-5,value6-Firmware Revision-6
            #             for i in range(2, NewRows1 + 1):
            #                 value1 = TRtemps2.Cells(2, 1).Value
            #                 value2 = TRtemps2.Cells(2, 2).Value
            #                 value3 = TRtemps2.Cells(2, 3).Value
            #                 if TRule12s.Cells(i, 2).Value == value1:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 2).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 2).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(2) + ' ' + 'correct value should be ' + value1
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 7).Value == value2:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 7).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 7).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(7) + ' ' + 'correct value should be ' + value2
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 8).Value == value3:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 8).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 8).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(8) + ' ' + 'correct value should be ' + value3
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 4).Value == None or TRule12s.Cells(i, 4).Value == 'N/A':
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 4).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 4).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(
            #                         4) + ' ' + 'correct value should be ' + 'None or N/A'
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 5).Value == None or TRule12s.Cells(i, 5).Value == 'N/A':
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 5).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 5).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(
            #                         5) + ' ' + 'correct value should be ' + 'None or N/A'
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 6).Value != None:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 6).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 6).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(6) + ' ' + 'correct value should not be ' + 'None'
            #                     logt.write(log + '\n')
            #                     # 实现Rule5
            #     else:
            #         # Localization Rule1&Rule6
            #         # Rule1
            #         TRule12s = TRule12.Sheets(Li)
            #         Rows1 = TRule12s.UsedRange.Rows.Count
            #         TRtemps = TRtemp.Sheets('Rule1')
            #         Rows2 = TRtemps.UsedRange.Rows.Count
            #         NewRows1 = Rows1
            #         for pd in range(1, Rows1 + 1):
            #             if not TRule12s.Cells(pd, 1).Value:
            #                 NewRows1 = pd - 1
            #                 break
            #             else:
            #                 pass
            #         print NewRows1
            #         if NewRows1 == 1:
            #             pass
            #         else:
            #             value = []
            #             for rows2 in range(2, Rows2 + 1):
            #                 if TRtemps.Cells(rows2, 1).Value == Li:
            #                     value.append(TRtemps.Cells(rows2, 2).Value)
            #                 else:
            #                     pass
            #             for i in range(2, NewRows1 + 1):
            #                 if TRule12s.Cells(i, 11).Value in value:
            #                     pass
            #                 else:
            #                     # 不符合rule,数据字体标红
            #                     # TRule12s.Cells(i, 11).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 11).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(11)
            #                     logt.write(log + '\n')
            #             # Rule6
            #             TRtemps2 = TRtemp.Sheets('Rule6')
            #             # value1-Responsible Manager-7,value2-Purchase Order-8
            #             # value3-Target Milestone-4,value4-Target Launch-5,value5-Firmware Revision-6
            #             for i in range(2, NewRows1 + 1):
            #                 value1 = TRtemps2.Cells(2, 1).Value
            #                 value2 = TRtemps2.Cells(2, 2).Value
            #                 value3 = TRtemps2.Cells(2, 3).Value

            #                 if TRule12s.Cells(i, 2).Value == value1:
            #                     pass
            #                 else:
            #                     TRule12s.Cells(i, 2).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(2) + ' ' + 'correct value should be ' + value1
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 7).Value == value2:
            #                     pass
            #                 else:
            #                     TRule12s.Cells(i, 7).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(7) + ' ' + 'correct value should be ' + value2
            #                     logt.write(log + '\n')
            #                 if TRule12s.Cells(i, 8).Value == value3:
            #                     pass
            #                 else:
            #                     TRule12s.Cells(i, 8).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(8) + ' ' + 'correct value should be ' + value3
            #                     logt.write(log + '\n')
            #                 # 2016/5/13 判断条件改为不能为N/A或None
            #                 if TRule12s.Cells(i, 4).Value != None and TRule12s.Cells(i, 4).Value != 'N/A':
            #                     pass
            #                 else:
            #                     TRule12s.Cells(i, 4).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(4)
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 5).Value != None and TRule12s.Cells(i, 4).Value != 'N/A':
            #                     pass
            #                 else:
            #                     TRule12s.Cells(i, 5).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(5)
            #                     logt.write(log + '\n')

            #                 if TRule12s.Cells(i, 6).Value != None:
            #                     pass
            #                 else:
            #                     # TRule12s.Cells(i, 6).Font.ColorIndex = 3
            #                     TRule12s.Cells(i, 6).Interior.ColorIndex = 3
            #                     logr = logr + 1
            #                     log = Li + ' ' + str(i) + ' ' + str(6)
            #                     logt.write(log + '\n')
            #                     # Rule6
            # Newfolderpath = folderpath.replace('/', "\\")
            # TRule12.SaveAs(Newfolderpath + '\\rule12.xlsx')
            
            # TRule12.Close()
            # winwbapp.Quit()
            
            #2016/5/12增加一步对Rules的判断
            # 临时表格tempwb
            winwbapp = win32com.client.Dispatch('Excel.Application')
            winwbapp.Visible = False
            tempwb = winwbapp.Workbooks.Open(folderpath + '/Temple.xlsx')
            TTwb = winwbapp.Workbooks.Open(folderpath + '/result.xlsx')
            TRtemp = winwbapp.Workbooks.Open(folderpath + '/Rules.xlsx')
            tempwbs = tempwb.Sheets(1)
            # m 表示TTtempS表格要写入数据的行数
            m = 0
            # 给定一个要合并数据的初始日期，实现从起始如期开始合并以后的所有数据
            # 创建一个变量Startdate,获取开始控件里的日期年 月 日
            gyear = variable1.get()
            gmonth = variable2.get()
            gday = variable3.get()
            Startdate = gmonth + '/' + gday + '/' + gyear
            # 添加一个结束日期Enddate,用来接收页面传来的日期
            yyear = variable4.get()
            mmonth = variable5.get()
            dday = variable6.get()
            Enddate = mmonth + '/' + dday + '/' + yyear
            # 抽取从开始日期到结束日期的所有数据
            for sheet in Z:
                print '********'
                print sheet
                print '********'
                # 读取结果表中前十二行的数据,遍历日期的列数range(12,Cols)分别找到开始日期和结束日期的下标
                SC = 0
                EC = 0
                TTwbs = TTwb.Sheets(sheet)
                Cols = TTwbs.UsedRange.Columns.Count
                Rows = TTwbs.UsedRange.Rows.Count
                for snc in range(13, Cols + 1):
                    if TTwbs.Cells(1, snc).Value == Startdate:
                        SC = snc
                    elif TTwbs.Cells(1, snc).Value == Enddate:
                        EC = snc
                    else:
                        pass
                # 2016/8/4------------------
                # for pd in range(1, Rows + 1):
                #     if TTwbs.Cells(pd, 1).Value == None:
                #         Rows = pd - 1
                #         break
                #     else:
                #         pass

                global blacklist 
                blacklist = []  
                if Rows == 1:
                    pass               
                else:
                    for j in range(SC, EC + 1):
                        for i in range(2, Rows + 1):
                            if i in blacklist:
                                continue
                            else: 
                                import decimal
                                if TTwbs.Cells(i, j).Value == None or TTwbs.Cells(i, j).Value == 0:
                                    for j2 in range(9,12):
                                        if TTwbs.Cells(i, j2).Value == 'Facility' :
                                            # print 'find facility'                  
                                            # 日期
                                            m = m+1 
                                            tempwbs.Cells(m + 1, 1).Value = Enddate
                                            # non offical charge 
                                            tempwbs.Cells(m + 1, 15).Value = TTwbs.Cells(i, j2 + 2 ).Value
                                            # 其他值
                                            for p in range(1,14):
                                                tempwbs.Cells(m + 1, p + 1).Value = TTwbs.Cells(i, p).Value

                                            blacklist.append(i)

                                        else:
                                            pass
                                else:
                                    m = m + 1
                                    # 把日期写入模板Temple的第 m+1行1列
                                    tempwbs.Cells(m + 1, 1).Value = TTwbs.Cells(1, j).Value
                                    # 把TTwbs.cell(i, j).value写入模板Temple的第m+1行index 14列
                                    tempwbs.Cells(m + 1, 14).Value = float(TTwbs.Cells(i, j).Value)
                                    for p in range(1, 13):
                                        # 2016/5/12增加一步判断,复制数据单元格背景颜色(WinSheet.Cells(1, 1).Interior.ColorIndex = 3)
                                        if TTwbs.Cells(i, p).Interior.ColorIndex == 3:
                                            tempwbs.Cells(m + 1, p + 1).Interior.ColorIndex = 3
                                            tempwbs.Cells(m + 1, p + 1).Value = TTwbs.Cells(i, p).Value
                                        else:
                                            tempwbs.Cells(m + 1, p + 1).Value = TTwbs.Cells(i, p).Value
                                            # 2016/5/12增加一步判断,复制数据单元格背景颜色(WinSheet.Cells(1, 1).Interior.ColorIndex = 3)
                
                
            global blacklist1
            blacklist1 = [] 
            Cols_temp = tempwbs.UsedRange.Columns.Count
            Rows_temp = tempwbs.UsedRange.Rows.Count
            refer_word1 = '[inOS]' 
            refer_word2 = '[HP ePrint]'
            for rows_temp in range (2, Rows_temp +1 ):
                if rows_temp in blacklist:
                    continue
                # 增加对各项值所对应的rule的判断
                elif tempwbs.cells(rows_temp,4).value == 'Test Execution - Auto':
                    if tempwbs.cells(rows_temp,9).value == 'HPI317341' :
                        blacklist1.append(rows_temp)
                    else:
                        tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                        logr = logr + 1
                        log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                        logt.write(log + '\n')
                        blacklist1.append(rows_temp)

                elif tempwbs.cells(rows_temp,3).value == 'Sys:Ink Sub' :
                    # if tempwbs.cells(rows_temp,4).value == 'Test Lead' or tempwbs.cells(rows_temp,4).value == 'Test Execution': 
                    if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                        blacklist1.append(rows_temp)
                    else:
                        tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                        logr = logr + 1
                        log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                        logt.write(log + '\n')
                        blacklist1.append(rows_temp)
                    

                elif tempwbs.cells(rows_temp,3).value == 'FW:Localization' :
                    if tempwbs.cells(rows_temp,4).value == 'Test Lead':
                        if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                           blacklist1.append(rows_temp)
                        else: 
                            tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                            logr = logr + 1
                            log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                            logt.write(log + '\n')
                            blacklist1.append(rows_temp) 
                    if tempwbs.cells(rows_temp,4).value == 'Test Execution':
                        if tempwbs.cells(rows_temp,2).value == 'Verona' or tempwbs.cells(rows_temp,2).value == 'Kronos Refresh WL' :
                            if tempwbs.cells(rows_temp,9).value == 'HPI217653-V3':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI217653-V3'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                        elif tempwbs.cells(rows_temp,2).value == 'Ellis High':
                            if tempwbs.cells(rows_temp,9).value == 'HPI218414':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI218414'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                        elif tempwbs.cells(rows_temp,2).value == 'Limo MFFS Hi':
                            if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                    if tempwbs.cells(rows_temp,4).value == 'Test Development':
                        if tempwbs.cells(rows_temp,12).value == 'Test Engineer I':
                            if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                        elif tempwbs.cells(rows_temp,12).value == 'Test Engineer II':
                            if tempwbs.cells(rows_temp,9).value == 'HPI217653-V3':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI217653-V3'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                        
                        elif tempwbs.cells(rows_temp,12).value == 'Test Engineer III - Special Rate':
                            if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                    if tempwbs.cells(rows_temp,4).value == 'Defect Management':
                        if tempwbs.cells(rows_temp,2).value == 'Limo MFFS Hi':
                            if tempwbs.cells(rows_temp,9).value == 'HPI317341':
                                 blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI317341'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)    
                        elif tempwbs.cells(rows_temp,2).value == 'Weber PDL':
                            if tempwbs.cells(rows_temp,9).value == 'HPI218414':
                                 blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI218414'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)    
                    
                elif tempwbs.cells(rows_temp,3).value == 'Sys:Mobile Print Apps':
                    if tempwbs.cells(rows_temp,4).value == 'Test Development' or tempwbs.cells(rows_temp,4).value == 'Test Execution' or tempwbs.cells(rows_temp,4).value == 'Test Lead':
                        if tempwbs.cells(rows_temp,9).value == 'HPI225388':
                            blacklist1.append(rows_temp)
                        else:
                            tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                            logr = logr + 1
                            log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI225388'
                            logt.write(log + '\n')
                            blacklist1.append(rows_temp)
                    else:
                        if refer_word1 in tempwbs.cells(rows_temp,13).value or refer_word2 in tempwbs.cells(rows_temp,13).value:
                                if tempwbs.cells(rows_temp,9).value == 'HPI225388':
                                    blacklist1.append(rows_temp)
                                else:
                                    tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                    logr = logr + 1
                                    log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI225388'
                                    logt.write(log + '\n')
                                    blacklist1.append(rows_temp)
                        else:
                            if tempwbs.cells(rows_temp,9).value == 'HPI217653-V3':
                                blacklist1.append(rows_temp)
                            else:
                                tempwbs.cells(rows_temp,9).Interior.ColorIndex = 3
                                logr = logr + 1
                                log = str(rows_temp) + ' ' + str(9) + ' ' + 'correct value should be ' + 'HPI217653-V3'
                                logt.write(log + '\n')
                                blacklist1.append(rows_temp)
                else:
                    pass

            logt.close()
            print logr                               
            Newfolderpath = folderpath.replace('/', "\\")
            tempwb.SaveAs(Newfolderpath + '\\test1.xlsx')
            TRtemp.Close()
            tempwb.Close()
            TTwb.Close()
            winwbapp.Quit()
            tempwbpath = folderpath + '/test1.xlsx'
            Newtempwbpath = tempwbpath.replace('/', "\\")
            winwbapp = win32com.client.Dispatch('Excel.Application')
            winwbapp.Visible = False
            tempwb = winwbapp.Workbooks.Open(Newtempwbpath)
            Ttempwb = winwbapp.Workbooks.Open( folderpath + '/Temple.xlsx')
            tempwbs = tempwb.Sheets(1)
            Ttempwbs = Ttempwb.Sheets(1)
            # 给定一个list f 存放tempwb中第一列日期数据
            f = []
            a = tempwbs.UsedRange.Rows.Count
            if a == 1:
                pass
            else:
                for aa in range(2, a + 1):
                    f.append(tempwbs.Cells(aa, 1).Value)
                # x表示过滤掉重复的日期
                x = set(f)
                # 将set类型转换成list类型
                xx = list(x)
                # 用sorted方法将日期list xx从小到大排序
                xxsorted = sorted(xx)
                # IT表示写入test2的数据的行数
                IT = 0
                c = 14
                for ydate in xxsorted:
                    for i in range(2, a + 1):
                        if tempwbs.Cells(i, 1).Value == ydate:
                            IT = IT + 1
                            for j in range(1, c + 1):
                                # 2016/5/12增加判断copy 单元格背景颜色
                                if tempwbs.Cells(i, j).Interior.ColorIndex == 3:
                                    Ttempwbs.Cells(IT + 1, j).Interior.ColorIndex = 3
                                    Ttempwbs.Cells(IT + 1, j).Value = tempwbs.Cells(i, j).Value
                                else:
                                    Ttempwbs.Cells(IT + 1, j).Value = tempwbs.Cells(i, j).Value
                                    # 2016/5/12增加判断copy 单元格背景颜色
                        else:
                            pass

            Ttempwbpath = folderpath + '/test2.xlsx'
            NewTtempwbpath = Ttempwbpath.replace('/', "\\")
            Ttempwb.SaveAs(NewTtempwbpath)
            Ttempwb.Close()
            tempwb.Close()
            winwbapp.Quit()
            # test2.xls的数据写入模板
            import win32com.client
            winapp = win32com.client.Dispatch('Excel.Application')
            winwbapp.Visible = False
            winwball = winapp.Workbooks.Open(folderpath + '/Template_ics_cbp_pactera-v6.0.xlsx')
            winwbtemp = winapp.Workbooks.Open(folderpath + '/test2.xlsx')
            tempS = winwbtemp.Sheets(1)
            # 2016/5/9增加一步判断，把数据写入CBP模板中的存放结果的sheet
            count = winwball.Sheets.Count
            R = 'ChargeByProjectDetails'
            for s in range(1, count + 1):
                allS = winwball.Sheets(s)
                if allS.Cells(1, 1).Value == 'Date':
                    R = allS.Name
                    print R
                    break
                else:
                    pass
            test2R = tempS.UsedRange.Rows.Count
            print test2R
            if test2R == 1:
                pass
            else:
                # 2016/5/9增加一步判断，把数据写入CBP模板中的存放结果的sheet
                allS = winwball.Sheets(R)
                # IT代表test2写入数据的长度,
                for i in range(2, test2R+1):
                    for j in range(1, 16):
                        # 2016/5/12增加一步判断copy 单元格的背景颜色(WinSheet.Cells(1, 1).Interior.ColorIndex = 3)
                        if tempS.Cells(i, j).Interior.ColorIndex == 3:
                            allS.Cells(i, j).Interior.ColorIndex = 3
                            allS.Cells(i, j).Value = tempS.Cells(i, j).Value
                        else:
                            allS.Cells(i, j).Value = tempS.Cells(i, j).Value
                            # 2016/5/12增加一步判断copy 单元格的背景颜色
                for gdr in range(2, test2R+1):
                    allS.Cells(gdr, 16).Value = 'Std Rate - Day Shift'
                for gdr in range(2, test2R+1):
                    allS.Cells(gdr, 17).Value = 'TAR'
                for i in range(2, test2R+1):
                    if allS.Cells(i, 3).Value == 'Other:Facility Costs':
                        allS.Cells(i, 15).Value = allS.Cells(i, 14).Value
                        allS.Cells(i, 14).Value = None
                    else:
                        pass
                # 2016/5/16 cbpf file format
                # SStartda开始日期 EEnddate结束日期
                gyear = variable1.get()
                gmonth = variable2.get()
                gmonthL = gmonth
                gmonthLL = len(gmonthL)
                if gmonthLL == 1:
                    gmonth = '0' + gmonth
                else:
                    pass
                gday = variable3.get()
                SStartdate = gyear + gmonth + gday
                yyear = variable4.get()
                mmonth = variable5.get()
                dday = variable6.get()
                mmonthL = mmonth
                mmonthLL = len(mmonthL)
                if mmonthLL == 1:
                    mmonth = '0' + mmonth
                else:
                    pass
                EEnddate = yyear + mmonth + dday
                cbpf = SStartdate + '_' + EEnddate + '_ics_cbp_pactera-v6.0.xlsx'
                # 2016/5/16 CBP file format
                winwball.SaveAs(Newfolderpath + '\\' + cbpf)
            winwbtemp.Close()
            winwball.Close()
            winapp.Quit()
            # Cenerate the CBP template success tips!
            # 2016/5/12增加一步判断,log的行数是否W为空,如果是空表示创建模板成功,如果为非空表示有错误
            # 2016/5/12删除多余文件
            # os.rename(os.path.join(folderpath, cbpf), os.path.join(folderpath, Newcbpf))
            filename1 = folderpath + '/rule12.xlsx'
            filename2 = folderpath + '/test1.xlsx'
            filename3 = folderpath + '/test2.xlsx'
            if os.path.exists(filename1):
                os.remove(filename1)
            if os.path.exists(filename2):
                os.remove(filename2)
            if os.path.exists(filename3):
                os.remove(filename3)
            if logr==0:
                tkMessageBox.showinfo("Creat successfully", "Generate the CBP template successfully,"
                                      + "please refer to " + folderpath + " for some details")
            else:
                tkMessageBox.showinfo("Creat successfully", "Some data does not inconsistent with rules," +
                                      "please refer to " + folderpath + "/log.txt for some details")
def callback3():
    sys.exit()
if __name__ == '__main__':
    # 布局以下
    import tkFileDialog
    import calendar

    root = Tk()
    root.title('CBP Results Summary')
    # root.geometry('550x160')
    lable1 = Label(root, text='FilePath:', font='Helvetica -12 bold')
    # entry1 应该设置为禁止输入的状态,state=DISABLED
    e1 = StringVar()
    entry1 = Entry(root, width=40,textvariable = e1)
    # entry1 = Entry(root)
    fram1 = Frame()
    button1 = Button(root, text='......', font='Helvetica -12 bold', command=callback1,width=4,height=1)
    lable2 = Label(root, text='Component:', font='Helvetica -12 bold')
    lable3 = Label(root, text='DateRange:', font='Helvetica -12 bold')
    lable4 = Label(fram1, text='-', font='Helvetica -12 bold')

    variable1 = StringVar(fram1)
    variable2 = StringVar(fram1)
    variable3 = StringVar(fram1)

    variable4 = StringVar(fram1)
    variable5 = StringVar(fram1)
    variable6 = StringVar(fram1)
    # 获取当前时间,年月日
    nowyear = calendar.datetime.datetime.now().year
    snowyear = str(nowyear)
    nowmonth = calendar.datetime.datetime.now().month
    snowmonth = str(nowmonth)
    nowday = calendar.datetime.datetime.now().day
    snowday = str(nowday)
    if calendar.datetime.datetime.now() is None:
        variable1.set('2016')  # default value
        variable2.set('6')  # default value
        variable3.set('1')  # default value
        variable4.set('2016')  # default value
        variable5.set('6')  # default value
        variable6.set('1')  # default value
    else:
        variable1.set(snowyear)
        variable2.set(snowmonth)
        variable3.set('1')
        variable4.set(snowyear)
        variable5.set(snowmonth)
        variable6.set(snowday)
    year = ['2016', '2017', '2018','2019', '2020']
    month = ['1', '2', '3','4', '5','6', '7', '8','9', '10','11', '12']
    day = []
    for days in range(1,32):
        kdays = str(days)
        day.append(kdays)
    combobox1 = ttk.Combobox(fram1,textvariable=variable1, values=year,font='Helvetica -12 bold',width='4')
    combobox2 = ttk.Combobox(fram1,textvariable=variable2, values=month,font='Helvetica -12 bold',width='2')
    combobox3 = ttk.Combobox(fram1,textvariable=variable3, values=day,font='Helvetica -12 bold',width='2')

    combobox4 = ttk.Combobox(fram1,textvariable=variable4, values=year,font='Helvetica -12 bold',width='4')
    combobox5 = ttk.Combobox(fram1,textvariable=variable5, values=month,font='Helvetica -12 bold',width='2')
    combobox6 = ttk.Combobox(fram1,textvariable=variable6, values=day,font='Helvetica -12 bold',width='2')
    button2 = Button(root, text='Creat', font='Helvetica -12 bold',command=callback2,width=4,height=1)
    button3 = Button(root, text='Quit', font='Helvetica -12 bold',command=callback3,width=4,height=1)

    lable1.grid(row=0, column=0, sticky=E)
    entry1.grid(row=0, column=1, columnspan=4, sticky=W)
    button1.grid(row=0, column=7, sticky=E)
    lable2.grid(row=1, column=0, sticky=E)
    lable3.grid(row=23, column=0, sticky=E)
    lable4.grid(row=23, column=0, sticky=E)
    fram1.grid(row=24, column=1, columnspan=7, rowspan=1,sticky=W)

    combobox1.grid(row=24, column=1,sticky=E)
    combobox2.grid(row=24, column=2,sticky=E)
    combobox3.grid(row=24, column=3,sticky=E)
    lable4.grid(row=24, column=4, sticky=E)
    combobox4.grid(row=24, column=5,sticky=E)
    combobox5.grid(row=24, column=6,sticky=E)
    combobox6.grid(row=24, column=7,sticky=E)
    button2.grid(row=25, column=6, sticky=E)
    button3.grid(row=25, column=7, sticky=E)
    # 以上布局以上
    root.mainloop()
    print '0k'



