#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlsxwriter
import os.path
import xlrd

#首先实例化一个xlsxwriter的Workbook()，这将创建一个Excel文件
workbook = xlsxwriter.Workbook('./汇总结果.xlsx')
#利用add_worksheet()方法添加一个工作簿
worksheet = workbook.add_worksheet()

# 粗体居中格式
#boold_center = workbook.add_format({'bold':True,'align':'center'})
# 写入标题
#worksheet.merge_range("A1:E1", "员工内购信息登记表",boold_center)

# 添加一个粗体格式
bold = workbook.add_format({'bold': True})
# 在Excel中写入项目名称
worksheet.write('A1',"序号",bold)
worksheet.write('B1',"本端设备名称",bold)
worksheet.write('C1',"命名",bold)
worksheet.write('D1',"本端设备管理IP",bold)
worksheet.write('E1',"本端设备所在机房模块",bold)
worksheet.write('F1',"起始端机架",bold)
worksheet.write('G1',"机房",bold)
worksheet.write('H1',"本端设备端口号",bold)
worksheet.write('I1',"对端设备名称",bold)
worksheet.write('J1',"命名",bold)
worksheet.write('K1',"对端设备管理IP",bold)
worksheet.write('L1',"对端设备所在机房模块",bold)
worksheet.write('M1',"目的端机架",bold)
worksheet.write('N1',"机房",bold)
worksheet.write('O1',"对端端口号",bold)

n = 2
for i in os.listdir('./'): # 循环查找目录下的各个excel文件
    # print(n)
    if i.startswith('~') is False and i.endswith('xlsx'):
        # print(i)
        file = xlrd.open_workbook(i) # 打开excel文件
        sheet_array = file.sheet_names() # 获取excel文件中的"sheet名称"数组
        for l in sheet_array: # 循环查找文件下的各sheet
            #info = file.sheet_by_index(0) # 通过索引获取sheet
            info = file.sheet_by_name(l) # 通过表名获取sheet
            rows = info.nrows # 获取sheet的总行数
            for m in range(1,rows): # 循环查找sheet内的各行数据
                num1 = info.cell(m,0).value # 序号
                devicename1 = info.cell(m,1).value # 本端设备名称
                name1 = info.cell(m,2).value # 命名
                ip1 = info.cell(m,3).value # 本端设备管理IP
                unum1 = info.cell(m,4).value # 本端设备所在机房模块
                cabinet1 = info.cell(m,5).value # 起始端机架
                address1 = info.cell(m,6).value # 机房
                port1 = info.cell(m,7).value # 本端设备端口号
                devicename2 = info.cell(m,8).value # 对端设备名称
                name2 = info.cell(m,9).value # 命名
                ip2 = info.cell(m,10).value # 对端设备管理IP
                unum2 = info.cell(m,11).value # 对端设备所在机房模块
                cabinet2 = info.cell(m,12).value # 目的端机架
                address2 = info.cell(m,13).value # 机房
                port2 = info.cell(m,14).value # 对端端口号
                worksheet.write("A{}".format(n),num1)
                worksheet.write("B{}".format(n),devicename1)
                worksheet.write("C{}".format(n),name1)
                worksheet.write("D{}".format(n),ip1)
                worksheet.write("E{}".format(n),unum1)
                worksheet.write("F{}".format(n),cabinet1)
                worksheet.write("G{}".format(n),address1)
                worksheet.write("H{}".format(n),port1)
                worksheet.write("I{}".format(n),devicename2)
                worksheet.write("J{}".format(n),name2)
                worksheet.write("K{}".format(n),ip2)
                worksheet.write("L{}".format(n),unum2)
                worksheet.write("M{}".format(n),cabinet2)
                worksheet.write("N{}".format(n),address2)
                worksheet.write("O{}".format(n),port2)
                n += 1
        print("完成{}的数据提取".format(i))
workbook.close()