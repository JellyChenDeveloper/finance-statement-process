import os

import pandas as pd

main_folder = '原始文件'
all_data = []
for root, dirs, files in os.walk(main_folder):
    for file in files:
        file_path = os.path.join(root, file)
        print(file_path)
        if file.startswith(".~"):
            continue
        if file.endswith(".xlsx"):  # 只处理Excel文件，可以根据需要修改扩展名
            df = pd.read_excel(file_path)[['仓库', '商品名称', '期末结存数量']]
            all_data.append(df)
        elif file.endswith(".xls"):
            file_path = os.path.join(root, file)
            df = pd.read_excel(file_path, engine='xlrd')[['仓库', '商品名称', '期末结存数量']]
            all_data.append(df)

combined_df = pd.concat(all_data, ignore_index=True)
s_1 = pd.isna(combined_df['仓库'])
combined_df = combined_df[s_1 == False]
combined_df.to_excel('合并后数据.xlsx', index=False)

# file = pd.ExcelFile('原始文件/10.30南疆.xls')
# names = file.sheet_names
# print(names)
#
# excel = pd.read_excel('10.30南疆.xls')
# # print(excel)
# print(excel.columns)
# # print(excel['仓库'])
# # print(excel['商品名称'excel])
# excel.drop(columns='规格型号', inplace=True)
# # print(excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库'])
# where = excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库']
# # excel.drop(index=where.index, inplace=True)
# excel.insert(2, column='总部商品名称', value='')
# excel.insert(3, column='总部商品型号', value='')
# compData = pd.read_excel('替换型号.xls')
# # print(compData.columns)
# # print(compData.values)
# print(excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库'].loc[excel['商品名称'] == '电烤箱KWS220-R015'])
# # 遍历每一行
# for index, row in compData.iterrows():
#     # print(f"Index: {index}, Row: {row['内部商品名称']}, {row['总部商品名称']}, {row['总部商品型号']}")
#     excel.loc[excel['商品名称'] == row['内部商品名称'], '总部商品名称'] = row['总部商品名称']
#     excel.loc[excel['商品名称'] == row['内部商品名称'], '总部商品型号'] = row['总部商品型号']
# print(excel.loc[excel['仓库'] == 'KA苏1库尔勒样机库'].loc[excel['商品名称'] == '电烤箱KWS220-R015'])
#
# # # 遍历每一列
# # for column, value in compData.items():
# #     print(f"Column: {column}")
# #     print(value)
# # for sheet, df in compData.items():
# # print(df)
# # for row in df.values:
# #     print(row)
#
