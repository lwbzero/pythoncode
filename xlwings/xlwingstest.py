#导入xlwings模块
import xlwings as xw

# 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
app=xw.App(visible=True,add_book=False)
# wb = app.books.active
# sht = wb.sheets.active
app.display_alerts=False
app.screen_updating=False
# wb=xw.books.active

# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
filepath=r'E:\\pythoncode\\pythoncode\\xlwings\\mac_data.xlsx'
wb=app.books.open(filepath)
# wb.sheets.add(name='woshisheet2',before=None,after=None)#添加一个新的工作表
wb.sheets['sheet1'].range('A1').value='人生'#给sheet1的A1写入value
# wb.sheets['woshisheet2'].range('A1').value='苦短'
# wb.sheets['sheet3'].range('A1').value='苦短'

# 示例代码
wb = xw.books['mac_data.xlsx']
wb = xw.books.active  #引用
sht = xw.books['mac_data.xlsx'].sheets['sheet1']
sht = xw.sheets.active
rng=xw.books['mac_data.xlsx'].sheets['sheet1'].range('A1')
height = 15
wb.sheets[0]['1:10'].api.RowHeight = 100 #设置行高100
# rng=xw.Range('A1')

#workbook api
# wbpath = wb.fullname
# print(wbpath)
# print(wb.name)

#sheet api
# sht = xw.books['workbook name'].sheets['sheet name']
# sht = sht.active()
# sht = sht.clear()
# sht = sht.contents()
# shtname = sht.name
# sht.delete


# rng = xw.Range('A1')
rng.resize(row_size=350,column_size=None)

wb.save()
wb.close()
app.quit()