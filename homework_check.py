
import os  
import sys
import time
import xlrd
import xlwt
import re
 

script_path = os.path.realpath(__file__)
script_dir = os.path.dirname(script_path)
script_name = os.path.basename(script_path)

#得到要统计的文件夹的路径和名字
homework_dir = sys.argv[1]
homework_name = os.path.basename(homework_dir)

print(homework_dir)
print(homework_name)


excel_file = '学号表.xlsx'
#i = 1
a = os.walk(homework_dir)
b = None 

#excel read (using sheet1)
wb = xlrd.open_workbook(filename=script_dir+"\\"+excel_file)#打开文件
print(wb.sheet_names())#获取所有表格名字
sheet1 = wb.sheet_by_index(0)#通过索引获取表格    
#sheet2 = wb.sheet_by_name('年级')#通过名字获取表格
#print(sheet1)
#print(sheet1.name,sheet1.nrows,sheet1.ncols)

#excel write (using sheet)
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
#创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
# 在电脑桌面右键新建一个Excel文件，其中就包含sheet1，sheet2，sheet3三张表
sheet = book.add_sheet('test', cell_overwrite_ok=True)
# 其中的test是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False
# 向表test中添加数据
#sheet.write(0, 0, 'englishname')  # 其中的'0-行, 0-列'指定表中的单元，'englishname'是向该单元写入的内容
#sheet.write(1, 0, 'marcovaldo')
#txt1 = '中文名字'
#sheet.write(0, 1, txt1.decode('utf-8'))  # 此处需要将中文字符串解码成unicode码，否则会报错
#txt2 = '马可瓦多'
#sheet.write(1, 1, txt2.decode('utf-8'))


#copy the first col of excel

cols = sheet1.col_values(0)#获取列内容，复制
i = 0
for sheet_data in cols:
    sheet.write(i, 0, label=sheet_data)
    i = i+1
#写统计列列名
sheet.write(0, 1, label=homework_name)

#统计作业情况

for root, dirs, files in os.walk(homework_dir):  
    #print(i)
    #i += 1
    #print(root) #当前目录路径  
    #print(dirs) #当前路径下所有子目录  
    print(files) #当前路径下所有非目录子文件 

    #rows = sheet1.row_values(2)#获取行内容
    #cols = sheet1.col_values(3)#获取列内容
    #print(rows)
    #print(cols)

    #print(sheet1.cell(1,0).value)#获取表格里的内容，三种方式
    #print(sheet1.cell_value(1,0))
    print(sheet1.row(0)[0].value)
    i = 0
    for sheet_data in cols:
        for fliename in files:
            pattern = sheet_data
            m = re.search(pattern, fliename)
            if m:
                sheet.write(i, 1, label='1')
        i = i+1

#print(b)
 
# 最后，将以上操作保存到指定的Excel文件中
book.save(script_dir+"\\"+r'统计结果.xls')  # 在字符串前加r，声明为raw字符串，这样就不会处理其中的转义了。否则，可能会报错
#time.sleep( 10 )
