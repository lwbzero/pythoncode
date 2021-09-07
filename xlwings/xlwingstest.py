#导入xlwings模块
from os import truncate
import xlwings as xw

# 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
app=xw.App(visible=True,add_book=False)
# wb = app.books.active
# sht = wb.sheets.active
app.display_alerts=False
app.screen_updating=False
# wb=xw.books.active

# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
filepath=r'E:\\pythoncode\\xlwings\\dangbei358.xlsx'
wb=app.books.open(filepath)
wb.sheets.add(name='woshisheet2',before=None,after=None)#添加一个新的工作表
wb.sheets['sheet1'].range('A1').value='人生'#给sheet1的A1写入value
wb.sheets['sheet2'].range('A1').value='苦短'
wb.save()
wb.close()
app.quit()

# 示例代码
# wb = xw.books['工作簿的名字']
# wb = xw.books.active  #引用
# sht = xw.books['工作簿的名字'].sheets['sheet的名字']
# sht = xw.sheets.active
# rng=xw.books['工作簿的名字'].sheets['sheet的名字'].range('A1')
# rng=xw.Range('A1')






