import os
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Map

# 1. 读取数据
file_path = 'AllData.xlsx'
if not os.path.exists(file_path):
    print(f"文件 {file_path} 不存在")
    exit()

# 映射字典，将Excel数据源中省份简写映射为Pyecharts需要的省份全称
province_mapping = {
    '河南': '河南省', 
    '广东': '广东省', 
    '河北': '河北省', 
    '云南': '云南省', 
    '四川': '四川省', 
    '北京': '北京市', 
    '山西': '山西省', 
    '山东': '山东省', 
    '贵州': '贵州省', 
    '广西': '广西壮族自治区', 
    '江苏': '江苏省', 
    '江西': '江西省', 
    '吉林': '吉林省', 
    '浙江': '浙江省', 
    '安徽': '安徽省', 
    '湖北': '湖北省', 
    '福建': '福建省', 
    '陕西': '陕西省', 
    '湖南': '湖南省', 
    '重庆': '重庆市', 
    '新疆': '新疆维吾尔自治区', 
    '上海': '上海市', 
    '黑龙江': '黑龙江省', 
    '辽宁': '辽宁省', 
    '天津': '天津市', 
    '内蒙古': '内蒙古自治区', 
    '香港': '香港特别行政区', 
    '台湾': '台湾省', 
    '宁夏': '宁夏回族自治区', 
    '青海': '青海省', 
    '甘肃': '甘肃省', 
    '海南': '海南省', 
    '西藏': '西藏自治区', 
    '澳门': '澳门特别行政区', 
    '-': '-', 
    '南海诸岛': '南海诸岛'
}

# 读取所有 sheet 数据
all_data = pd.read_excel(file_path, sheet_name=None)

# 通过映射字典更新省份列
for sheet_name, sheet_data in all_data.items():
    sheet_data['province'] = sheet_data['province'].map(province_mapping)

# 2. 统计每个 province 的总数量
province_counts = {}
for sheet_name, sheet_data in all_data.items():
    for index, row in sheet_data.iterrows():
        province = row['province']
        if pd.notna(province):
            province_counts[province] = province_counts.get(province, 0) + 1

# 3. 绘制地图
province_list = list(province_counts.keys())
count_list = list(province_counts.values())

map_chart = (
    Map()
    .add("访问量", [list(z) for z in zip(province_list, count_list)], "china")
    .set_series_opts(
        label_opts=opts.LabelOpts(is_show=True, formatter="{c}")
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(title="中国省份用户访问量分布"),
        visualmap_opts=opts.VisualMapOpts(max_=max(count_list)),
    )
)

# 4. 保存地图
map_chart.render("LearnPyecharts01.html")
