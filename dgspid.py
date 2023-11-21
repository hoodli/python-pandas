import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font

# 读取商品渠道数据
df = pd.read_excel('./商品渠道数据_2023-11-16.xlsx')
# df = pd.read_excel("D:\桌面\爆款渠道数据_2023-11-21.xlsx")
# 将商品ID列的格式转换为数字
df['商品ID'] = pd.to_numeric(df['商品ID'], errors='coerce')

# 添加商品ID条件筛选
product_ids = [566538659598, 527016966632, 526945915188,
               526983618990, 627856252682, 527028840637, 526994417188, 630682246162]
df_filtered = df[df['商品ID'].isin(product_ids)]

# 添加三级流量来源条件筛选
source_filter2 = ['淘内待分类', '我的淘宝', '购物车', '直通车',
                  '品销宝-品牌专区', '手淘淘宝直播', '手淘推荐', '手淘搜索', '淘宝客']
df_filtered2 = df_filtered[df_filtered['三级流量来源'].isin(source_filter2)]

# 根据“三级流量来源”列和统计日期列计算访客数、支付金额和下单买家数，并按照商品ID、三级流量来源和统计日期分组
grouped_data2 = df_filtered2.groupby(['商品ID', '三级流量来源', '统计日期']).agg({
    '访客数': 'sum',
    '支付金额': 'sum',
    '支付买家数': 'sum',
    '支付件数': 'sum'
})

# 计算pct(支付金额/下单买家数)并取整处理
grouped_data2['pct'] = grouped_data2['支付金额'] / grouped_data2['支付买家数']
grouped_data2['pct'] = grouped_data2['pct'].apply(
    lambda x: np.round(x) if pd.notnull(x) else x)

# 计算CRT(支付买家数/访客数)并以百分比形式输出
grouped_data2['CRT'] = (grouped_data2['支付买家数'] / grouped_data2['访客数'])


# 遍历每个商品ID，创建并保存工作簿
output_folder = r'/Users/kinxiaolei/Desktop/python'
# output_folder = r'D:\桌面\日报\爆款渠道数据'
for product_id in product_ids:
    workbook = Workbook()
    sheet = workbook.active
    data = grouped_data2.loc[product_id]
    data.reset_index(inplace=True)
    sheet.append(data.columns.tolist())
    for row in data.itertuples(index=False):
        sheet.append(row)

    # 设置字体为"微软雅黑"
    font = Font(name='微软雅黑')
    for column_cells in sheet.columns:
        for cell in column_cells:
            cell.font = font

    # 保存工作簿
    file_name = f"{output_folder}/{str(product_id)}.xlsx"
    workbook.save(file_name)
    workbook.close()
