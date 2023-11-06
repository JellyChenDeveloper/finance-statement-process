import os

import pandas as pd

main_folder = '委托四表'
all_data = []
for root, dirs, files in os.walk(main_folder):
    for file in files:
        file_path = os.path.join(root, file)
        print(file_path)
        if file.startswith(".~"):
            continue
        if file.endswith(".xlsx"):  # 只处理Excel文件，可以根据需要修改扩展名
            df = pd.read_excel(file_path, sheet_name=['委托'])
            all_data.append(df['委托'])
        elif file.endswith(".xls"):
            file_path = os.path.join(root, file)
            df = pd.read_excel(file_path, engine='xlrd', sheet_name=['委托'])
            all_data.append(df['委托'])

if len(all_data) == 0:
    print("未读取到数据")
elif len(all_data) > 1:
    combined_df = pd.concat(all_data, ignore_index=True)
else:
    combined_df = all_data[0]

print("Original data types:\n", combined_df['产品型号'].dtypes)
combined_df = combined_df[['产品型号', '结存数量']]
# combined_df['产品型号'] = combined_df['产品型号'].astype(str)
combined_df = combined_df.loc[combined_df['结存数量'].apply(lambda x: x > 0)]
# 获取当前计算出来的数据
originData = combined_df.groupby('产品型号', as_index=False)['结存数量'].sum()
# TODO 从委托四表中获取到的产品型号为字符串，但是在原始表中获取到的为数字和字符串混合模式
for i in originData.index:
    originData.at[i, '产品型号'] = pd.to_numeric(originData.at[i, '产品型号'], errors='ignore')
# originData['产品型号'] = pd.to_numeric(originData['产品型号'], errors='ignore')
print(originData)

template_df = pd.read_excel('样表/零售账存差异表单.xlsx', sheet_name='零售', skiprows=9, header=None)
del_data = template_df.iloc[:, [2, 19]]
del_data = del_data.rename(columns={2: '产品型号', 19: '零售受托数量'})
del_data = del_data[del_data['产品型号'].notnull()]
# del_data['产品型号'] = del_data['产品型号'].astype(str)
# del_data.set_index('产品型号', inplace=True, drop=False)
for xh in del_data['产品型号']:
    if xh in originData['产品型号']:
        del_data.loc[del_data['产品型号'] == xh, '零售受托数量'] = originData.loc[originData['产品型号'] == xh]
    else:
        del_data.loc[del_data['产品型号'] == xh, '零售受托数量'] = 2
print(del_data)
# for xh in originData.index:
#     if xh not in del_data['产品型号']:
#         del_data = del_data._append(pd.Series({'产品型号': xh, '零售受托数量': originData[xh]}), ignore_index=True)
# print("Original data types:\n", del_data['产品型号'].dtypes)
# print(del_data)
del_data.to_excel('生成的零售受托数量.xlsx', index=False)
# template_df.columns
# workbook = load_workbook('样表/零售账存差异表单.xlsx')
# worksheet = workbook['零售']
# worksheet['T10'] = 11111
# workbook.save('生成的零售账存差异表单.xlsx')
