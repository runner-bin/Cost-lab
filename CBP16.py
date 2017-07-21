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
            winwbapp = win32com.client.Dispatch('Excel.Application')
            winwbapp.Visible = False
            TRule12 = winwbapp.Workbooks.Open(folderpath + '/result.xlsx')
            TRtemp = winwbapp.Workbooks.Open(folderpath + '/Rule.xlsx')
            # 字典添加序列及对应行
            
            resultcell = {}
            resultcelldi={}
            rulecelldi ={}
            rulecell = []
            comparisoncell = []
            Resultblank = []
            global Resultblank


            def addWord(theIndex,word,word_value): 
                theIndex.setdefault(word, [ ]).append(word_value)




            
            for Li in Z:
                Rulerow = TRtemp.Sheets(Li).UsedRange.Rows.Count
                Resultrow = TRule12.Sheets(Li).UsedRange.Rows.Count  
                # 判断值必须在规则中 
                for sever in [1,2,3,4,5,7,10,11]: 
                    for Resultsever in range(2,Resultrow+1):
                        if TRule12.Sheets(Li).Cells(Resultsever,sever).value != None:
                            word = TRule12.Sheets(Li).Cells(Resultsever,sever).value
                            addWord(resultcell,word,Resultsever)
                    # 对sd与inksub特殊判断
                    if Li == 'SD&Others' or Li == 'Ink Sub':
                        for Rulesever in range (13,18):
                            if TRtemp.Sheets(Li).Cells(Rulesever,sever).value != None:
                                rulecell.append(TRtemp.Sheets(Li).Cells(Rulesever,sever).value)
                    # 对普通表格的判断
                    else:
                        for Rulesever in range(13, Rulerow+1):
                            if TRtemp.Sheets(Li).Cells(Rulesever,sever).value != None:
                                rulecell.append(TRtemp.Sheets(Li).Cells(Rulesever,sever).value)  
                    for key in resultcell:
                        if not key in rulecell:
                            comparisoncell =(resultcell.get(key))
                            for comparisever in comparisoncell:
                                TRule12.Sheets(Li).Cells(comparisever,sever).Interior.ColorIndex = 3
                                logr = logr + 1
                                log =  'No found in rule:' + Li + ' ' + str(comparisever) + ' ' + str(sever)
                                logt.write(log + '\n')
                            comparisoncell = []
                        else:
                            pass
                    dellist = {}
                    resultcell = dellist

                # 进一步判断
                # sd,1, 判断值不为空（1st Check point）
                if Li == 'SD&Others':
                    for Resultsever in range(2,Resultrow+1):
                        for ResultRowsever in [9,12]:
                            if TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).value == '':
                                TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).Interior.ColorIndex =3
                                logr = logr + 1
                                log =  'Data is blank:' + Li + ' ' + str(comparisever) + ' ' + str(sever)
                                logt.write(log + '\n')
                        Resultblank = []
                        for ResultRowsever in [14,15,16,17,18]:
                            if TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).value != None:
                                Resultblank.append(Resultsever)
 
                    if not Resultsever in Resultblank:
                        # for ResultRowsever in [14,15,16,17,18]:
                        #     TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).Interior.ColorIndex =3
                        #     Resultseverr = Resultsever
                        logr = logr + 1
                        log =  'Data is blank:' + Li + ' ' + ' line: '+str(Resultsever)
                        logt.write(log + '\n')
                # sd,2, 判断特殊条件 （2nd Check Point）
                    for Rulesever in range(20,Rulerow+1):
                        if TRtemp.Sheets(Li).Cells(Rulesever,13).value != None:
                            if TRtemp.Sheets(Li).Cells(Rulesever,12).value in str(TRule12.Sheets(Li).Cells(Resultsever,12).value)\
                            and TRtemp.Sheets(Li).Cells(Rulesever,13).value in str(TRule12.Sheets(Li).Cells(Resultsever,12).value):                       
                                for Ruleseverr in range(1,12):
                                    if TRtemp.Sheets(Li).Cells(Rulesever,Ruleseverr)!='':
                                        if TRule12.Sheets(Li).Cells(Resultsever,Ruleseverr).value == TRtemp.Sheets(Li).Cells(Rulesever,Ruleseverr).value:
                                            pass
                                        else:
                                            TRule12.Sheets(Li).Cells(Resultsever,Ruleseverr).Interior.ColorIndex =3
                                            logr = logr + 1
                                            log =  'The data is not correspond to rule:' + Li + ' ' + str(Resultsever) + ' ' + str(Ruleseverr)
                                            logt.write(log + '\n')    
                        else:
                            for Resultsever in range(2,Resultrow+1):
                                if TRtemp.Sheets(Li).Cells(Rulesever,12).value in TRule12.Sheets(Li).Cells(Resultsever,12).value:
                                    for Ruleseverr in range(1,12):
                                        if TRtemp.Sheets(Li).Cells(Rulesever,Ruleseverr)!='':
                                            if TRule12.Sheets(Li).Cells(Resultsever,Ruleseverr).value == TRtemp.Sheets(Li).Cells(Rulesever,Ruleseverr).value:
                                                pass
                                            else:
                                                TRule12.Sheets(Li).Cells(Resultsever,Ruleseverr).Interior.ColorIndex =3
                                                logr = logr + 1
                                                log =  'The data is not correspond to rule:' + Li + ' ' + str(Resultsever) + ' ' + str(Ruleseverr)
                                                logt.write(log + '\n')                 
                # 普通表格，判断值不为空 (1st Check Below first)
                else:



                    if not Resultsever in range(2,Resultrow+1):
                        for ResultRowsever in [6,9,12]:
                            if TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).value == '':
                                TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).Interior.ColorIndex =3
                                logr = logr + 1
                                log =  'The data is blank:' + Li + ' ' + str(ResultRowsever) + ' ' + str(sever)
                                logt.write(log + '\n')
                        Resultblank = []
                        for ResultRowsever in [16,17,18,19,20]:
                            if TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).value != None:
                                Resultblank.append(Resultsever)
                        for Resultsever in Resultblank:
                            # for ResultRowsever in [16,17,18,19,20]:
                            #     TRule12.Sheets(Li).Cells(Resultsever,ResultRowsever).Interior.ColorIndex =3
                            #     Resultseverr = Resultsever
                            logr = logr + 1
                            log =  'Data is blank:' + Li + ' ' + ' line: '+str(Resultsever)
                            logt.write(log + '\n')
                # 针对inksub team特殊规则(2nd Check Point)
                if Li == 'Ink Sub':
                    for Rulesever in range (19,Rulerow+1):
                        Rulelistsever = [TRtemp.Sheets(Li).Cells(Rulesever,3).value, TRtemp.Sheets(Li).Cells(Rulesever,10).value ]
                        word = TRtemp.Sheets(Li).Cells(Rulesever,11).value
                        addWord(rulecelldi,word,Rulelistsever)
                    for Resultsever in range(2,Resultrow+1):
                        Rusultlistsever = [TRule12.Sheets(Li).Cells(Resultsever,3).value, TRule12.Sheets(Li).Cells(Resultsever,10).value ]
                        addWord(rulecelldi,word,Rulelistsever)
                    for key in resultcelldi:
                        if key in rulecelldi:
                            if resultcelldi.get(key) == rulecelldi.get(key):
                                pass
                            else:
                                logr = logr + 1
                                log =  'No found this data' + Li + ' ' + str(resultcelldi.get(key))
                                logt.write(log + '\n')                             
                        else:
                            logr = logr + 1
                            log =  'No found this data' + Li + ' ' + str(resultcelldi.get(key))
                            logt.write(log + '\n')  
                # if Li == 'SD&Others': 
                #     pass
                # else:
                    
                #     if Resultrow == 2:
                #         pass
                #     else:  
                        
                #         column_in = [1]

                        # ,2,3,4,5,7,10,11
                        # for Resultsever in range(2, Resultrow+1):
                        #     print 'error code2'
                        #     for columnsever in column_in:
                        #         print 'error code3'
                        #         for Rulesever in range(13, Rulerow+1):
                        #             print 'error code4'
                        #             if TRtemp.Sheets('IMF').Cells(Rulesever, 1).value == '':
                        #                 print 'error code5'

                        #             else:
                        #                 if 'Limo MFFS Hi' == TRtemp.Sheets('IMF').Cells(Rulesever, 1).value:
                        #                     print 'pass'
                        #                     print TRule12.Sheets('IMF').Cells(Resultsever,1).value
                        #                     break
                        #                 else:
                        #                     print 'error code6'


                

                                

            # Resultrowtt = TRtemp.Sheets('IMF').UsedRange.Rows.Count   
                
        
            # print Resultrowtt
            # ddd = TRtemp.Sheets('Auto DEV').Cells(2,2).value
         

    
                    
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
            Newfolderpath = folderpath.replace('/', "\\")
            TRule12.SaveAs(Newfolderpath + '\\rule12.xlsx')
            TRtemp.close
            TRule12.Close()
            winwbapp.Quit()
            
            #2016/5/12增加一步对Rules的判断
            # 临时表格tempwb
            winwbapp = win32com.client.Dispatch('Excel.Application')
            winwbapp.Visible = False
            tempwb = winwbapp.Workbooks.Open(folderpath + '/Temple.xlsx')
            TTwb = winwbapp.Workbooks.Open(folderpath + '/rule12.xlsx')
            # TRtemp = winwbapp.Workbooks.Open(folderpath + '/Rules.xlsx')
            tempwbs = tempwb.Sheets(1)
            tempwbs2 = tempwb.Sheets(2)
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
                                            m = m+1 
                                            tempwbs2.Cells(m + 1, 1).Value = Enddate
                                            tempwbs2.Cells(m + 1, 15).Value = TTwbs.Cells(i, j2 + 2 ).Value
                                            # 其他值
                                            for p in range(1,14):
                                                tempwbs2.Cells(m + 1, p + 1).Value = TTwbs.Cells(i, p).Value
                                            blacklist.append(i)
                                        else:
                                            pass
                                else:
                                    m = m + 1
                                    # 把日期写入模板Temple.sheet(2)的第 m+1行1列
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
            # print logr                               
            Newfolderpath = folderpath.replace('/', "\\")
            tempwb.SaveAs(Newfolderpath + '\\test1.xlsx')
            # TRtemp.Close()
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
            tempwbs2 = tempwb.Sheets(2) 
            Ttempwbs = Ttempwb.Sheets(1)
            Ttempwbs2 = Ttempwb.Sheets(2)
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
                rowt1 = Ttempwbs.UsedRange.Rows.Count
                rowt2 = tempwbs2.UsedRange.Rows.Count
                newrow = rowt1+1
                sanda = 0
                for row2 in range(2,rowt2+1):
                    for j in range(1,16):
                        if tempwbs2.Cells(row2, j).Interior.ColorIndex == 3:
                            Ttempwbs.Cells(newrow +sanda,j ).Interior.ColorIndex == 3
                            Ttempwbs.Cells(newrow +sandas,j ).value = tempwbs2.Cells(row2, j).value
                            sanda +=1
                        else:
                            Ttempwbs.Cells(newrow,j ).value = tempwbs2.Cells(row2, j).value
                            sanda +=1
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
            winwbrp = winapp.Workbooks.Open(folderpath + '/Rule.xlsx')
            winwbrp_1 = winwbrp.Sheets(1)
            winwbrp_2 = winwbrp.Sheets(2)
            winwbrp_3 = winwbrp.Sheets(3)
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
                # PO规则的判断
                global blacklist1
                global smart
                smart=0
                blacklist1 = [] 
                for i in range (2, test2R +1 ):
                    if i in blacklist:
                        continue
                    # 增加对各项值所对应的rule的判断
                    elif allS.cells(i,3).value == 'Sys:Ink Sub' :
                        if allS.cells(i,9).value == winwbrp_3.Cells(2,9).value:
                            blacklist1.append(i)
                        else:
                            allS.cells(i,9).Interior.ColorIndex = 3
                            logr = logr + 1
                            log = '********** BELOW IS PO PROBLEM **********' 
                            logt.write(log + '\n')
                            logr = logr + 1
                            log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_3.Cells(2,9).value
                            logt.write(log + '\n')
                            blacklist1.append(i)
                            smart += 1
                    
                    elif allS.cells(i,3).value == 'Sys:Mobile Print Apps':
                        row_m = winwbrp_2.UsedRange.Rows.Count
                        for im in range(2,row_m +1):
                            if winwbrp_2.Cells(im,13) != "": 
                                if str(winwbrp_2.Cells(im,13).value) in str(allS.cells(i,13).value):
                                        if allS.cells(i,9).value == winwbrp_2.Cells(im,9).value:
                                            blacklist1.append(i)
                                        else:
                                            if smart == 0:
                                                logr = logr + 1
                                                log = '********** BELOW IS PO PROBLEM **********' 
                                                logt.write(log + '\n')                                              
                                            else:
                                                pass                                             
                                            allS.cells(i,9).Interior.ColorIndex = 3
                                            logr = logr + 1
                                            log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_2.Cells(im,9).value
                                            logt.write(log + '\n')
                                            blacklist1.append(i)
                                            smart 
                            else:
                                if allS.cells(i,9).value == winwbrp_2.Cells(im,9).value :
                                    blacklist1.append(i)
                                else:
                                    if smart == 0:
                                        logr = logr + 1
                                        log = '********** BELOW IS PO PROBLEM **********' 
                                        logt.write(log + '\n')                                              
                                    else:
                                        pass   
                                    allS.cells(i,9).Interior.ColorIndex = 3
                                    logr = logr + 1
                                    log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_2.Cells(im,9).value
                                    logt.write(log + '\n')
                                    blacklist1.append(i)

                    elif allS.cells(i,3).value == 'FW:Localization' :
                        row_l = winwbrp_1.UsedRange.Rows.Count
                        for il in range(2,row_l +1):
                            if  winwbrp_1.Cells(il, 2)!= "":
                                if winwbrp_1.Cells(il,2).value == allS.cells(i,2).value: 
                                    if allS.cells(i,9).value == winwbrp_1.Cells(il,9).value:
                                        blacklist1.append(i)
                                        # print '1-1'
                                    else:
                                        if smart == 0:
                                            logr = logr + 1
                                            log = '********** BELOW IS PO PROBLEM **********' 
                                            logt.write(log + '\n')                                              
                                        else:
                                            pass   
                                        allS.cells(i,9).Interior.ColorIndex = 3
                                        logr = logr + 1
                                        log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_1.Cells(il,9).value
                                        logt.write(log + '\n')
                                        blacklist1.append(i)
                                        # print '1-2'
                        
                            elif winwbrp_1.Cells(il,4) != "" :
                                if winwbrp_1.Cells(il,4).value == allS.cells(i,4).value: 
                                    if allS.cells(i,9).value == winwbrp_1.Cells(il,9).value:
                                        blacklist1.append(i)
                                        # print '2-1' 
                                    else:
                                        if smart == 0:
                                            logr = logr + 1
                                            log = '********** BELOW IS PO PROBLEM **********' 
                                            logt.write(log + '\n')                                              
                                        else:
                                            pass 
                                        allS.cells(i,9).Interior.ColorIndex = 3
                                        logr = logr + 1
                                        log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_1.Cells(il,9).value
                                        logt.write(log + '\n')
                                        blacklist1.append(i)
                                        
                            elif winwbrp_1.Cells(il,13)!="":
                                if str(winwbrp_1.Cells(il,13).value) in str(allS.cells(i,13).value): 
                                    if allS.cells(i,9).value == winwbrp_1.Cells(il,9).value:
                                        blacklist1.append(i)
                                        
                                    else:
                                        if smart == 0:
                                            logr = logr + 1
                                            log = '********** BELOW IS PO PROBLEM **********' 
                                            logt.write(log + '\n')                                              
                                        else:
                                            pass 
                                        allS.cells(i,9).Interior.ColorIndex = 3
                                        logr = logr + 1
                                        log = str(i) + ' ' + str(9) + ' ' + 'correct value should be ' + winwbrp_1.Cells(il,9).value
                                        logt.write(log + '\n')
                                        blacklist1.append(i) 
                # 判读milestone与target launch相符合，如果都是na、都不是na,否则log
                global smartT
                smartT = 0
                for i in range (2, test2R +1 ):
                    if allS.cells(i,5).value =='N/A' and allS.cells(i,6).value =='N/A':
                        pass
                    elif allS.cells(i,5).value=='NA' and allS.cells(i,6).value =='N/A':
                        pass
                    elif allS.cells(i,5).value=='N/A' and allS.cells(i,6).value =='NA':
                        pass
                    elif allS.cells(i,5).value !='N/A' and allS.cells(i,6).value !='N/A':
                        pass
                    else:
                        if smartT == 0:
                            logr = logr + 1
                            log = '********** BELOW IS "Target Milestone" & "Target Launch"  PROBLEM **********' 
                            logt.write(log + '\n')                                         
                        else:
                            pass                        
                        allS.cells(i,5).Interior.ColorIndex = 3 
                        allS.cells(i,6).Interior.ColorIndex = 3 
                        logr = logr + 1
                        log = "line" + str(i) + '  ' + str(5) + ' and ' + str(6) + " are not correspond"
                        logt.write(log + '\n')
                        print "4"
                        smartT +=1


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
            logt.close()    
            winwbtemp.Close()
            winwball.Close()
            winwbrp.Close()
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



