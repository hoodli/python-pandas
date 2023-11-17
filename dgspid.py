import pandas as pd
import numpy as np

#脚本解决从渠道数据中按照商品ID筛选出数据，并根据三级数据来源统计访客数、支付金额、支付买家数，计算出pct(支付金额/下单买家数)，CRT(支付买家数/访客数)并将结果输出到不同的excel表格，以商品ID命名。

# 读取商品渠道数据
df = pd.read_excel('./商品渠道数据_2023-11-16.xlsx')

# 将商品ID列的格式转换为数字
df['商品ID'] = pd.to_numeric(df['商品ID'], errors='coerce')

# 添加商品ID条件筛选
product_ids = [566538659598, 527016966632, 526945915188,
               526983618990, 627856252682, 527028840637]
df_filtered = df[df['商品ID'].isin(product_ids)]

# 添加三级流量来源条件筛选
source_filter2 = ['淘内待分类', '我的淘宝', '购物车', '直通车', '品销宝-品牌专区', '手淘淘宝直播', '淘宝客']
df_filtered2 = df_filtered[df_filtered['三级流量来源'].isin(source_filter2)]

# 根据“三级流量来源”列和统计日期列计算访客数、支付金额和下单买家数，并按照商品ID、三级流量来源和统计日期分组
grouped_data2 = df_filtered2.groupby(['商品ID', '三级流量来源', '统计日期']).agg({
    '访客数': 'sum',
    '支付金额': 'sum',
    '支付买家数': 'sum'
})

# 计算pct(支付金额/下单买家数)并取整处理
grouped_data2['pct'] = grouped_data2['支付金额'] / grouped_data2['支付买家数']
grouped_data2['pct'] = grouped_data2['pct'].apply(
    lambda x: np.round(x) if pd.notnull(x) else x)

# 计算CRT(支付买家数/访客数)并以百分比形式输出
grouped_data2['CRT'] = (grouped_data2['支付买家数'] / grouped_data2['访客数']) * 100

# 将数据保存到多个Excel表格
for product_id in product_ids:
    sheet_name = str(product_id) + '-三级'
    file_name = str(product_id) + '.xlsx'
    data = grouped_data2.loc[product_id]
    data.to_excel(file_name, sheet_name=sheet_name, index=True)
