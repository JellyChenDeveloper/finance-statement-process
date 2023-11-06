import pandas as pd

file = pd.ExcelFile('原始文件/10.30南疆.xls')
names = file.sheet_names
print(names)

excel = pd.read_excel('10.30南疆.xls')
# print(excel)
print(excel.columns)
# print(excel['仓库'])
# print(excel['商品名称'excel])
excel.drop(columns='规格型号', inplace=True)
# print(excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库'])
where = excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库']
# excel.drop(index=where.index, inplace=True)
excel.insert(2, column='总部商品名称', value='')
excel.insert(3, column='总部商品型号', value='')
compData = pd.read_excel('替换型号.xls')
# print(compData.columns)
# print(compData.values)
# 遍历每一行
for index, row in compData.iterrows():
    # print(f"Index: {index}, Row: {row['内部商品名称']}, {row['总部商品名称']}, {row['总部商品型号']}")
    excel.loc[excel['商品名称'] == row['内部商品名称'], '总部商品名称'] = row['总部商品名称']
    excel.loc[excel['商品名称'] == row['内部商品名称'], '总部商品型号'] = row['总部商品型号']
print(excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库'].loc[excel['商品名称'] == '电烤箱KWS220-R015'])

# # 遍历每一列
# for column, value in compData.items():
#     print(f"Column: {column}")
#     print(value)
# for sheet, df in compData.items():
# print(df)
# for row in df.values:
#     print(row)

# header = ('姓名', '年龄')
# rows = [('张三', 20), ('李四', 25)]
# df = pd.DataFrame(rows, columns=header)
# df.to_excel('test.xlsx', sheet_name='Sheet1', index=False)
