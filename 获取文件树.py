import os
import re
path= os.getcwd()
#用之前pip install xlwt
import xlwt
import datetime
from openpyxl import Workbook
workbook=xlwt.Workbook(encoding='utf-8')
sheet=workbook.add_sheet(path[0],cell_overwrite_ok=True)#解决单元格重复写入报错问题
m=0
x=0
for dirpath, next_dirnames, files in os.walk(path):
  dirname=dirpath[dirpath.rfind('\\')+1:]#获取文件夹名称和绝对路径：取文件夹绝对路径最后一个斜杠后面的名称
  n=dirpath.count("\\")-1
  # print('\t'*n,dirname,'dir:%s' % dirpath)
  link='HYPERLINK("%s";"%s")'%(dirpath,dirname)
  style = xlwt.XFStyle() # 初始化样式
  font = xlwt.Font() # 为样式创建字体
  font.name = 'Times New Roman' 
  font.bold = True # 黑体
  font.underline = True # 下划线
  font.italic = True # 斜体字
  font.colour_index = 2#文件夹的字体颜色设置为红色
  style.font = font # 设定样式
  sheet.write(m+x,n,xlwt.Formula('%s'%link),style)#行 列 写入的值 0代表第一
  x+=1
  for file in files:
    # print(f'\t'*(n+2),'%s' % file,os.path.abspath(file))#获取文件名称和绝对路径：取文件名称，\t代表一个tab
    filepath=dirpath+'\\'+file#不用os.path.abspath()的原因是它获取的不是真实路径http://www.imooc.com/wenda/detail/556337
    link='HYPERLINK("%s";"%s")'%(filepath,file)
    sheet.write(m+x,n+1,xlwt.Formula('%s'%link))#行 列 写入的值 0代表第一
    # sheet.write(m+x,n+1,file)
    m+=1
workbook.save("汇总登记"+'.xlsx')#保存到当前文件夹