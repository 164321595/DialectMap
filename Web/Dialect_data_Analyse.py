import pandas as pd
import json
import os
import random
from pathlib import Path
import re

def extract_province(location):
    """从地点名中提取省份信息"""
    if not location or pd.isna(location):
        return "未知省份"
    
    # 中国省份列表（包含直辖市、自治区、特别行政区）
    provinces = [
        "北京市", "天津市", "河北省", "山西省", "内蒙古自治区",
        "辽宁省", "吉林省", "黑龙江省", "上海市", "江苏省",
        "浙江省", "安徽省", "福建省", "江西省", "山东省",
        "河南省", "湖北省", "湖南省", "广东省", "广西壮族自治区",
        "海南省", "重庆市", "四川省", "贵州省", "云南省",
        "西藏自治区", "陕西省", "甘肃省", "青海省", "宁夏回族自治区",
        "新疆维吾尔自治区", "台湾省", "香港特别行政区", "澳门特别行政区"
    ]
    
    # 先尝试完全匹配
    for province in provinces:
        if province in location:
            return province
    
    # 再尝试匹配省份简称
    province_abbr = {
        "京": "北京市", "津": "天津市", "冀": "河北省", "晋": "山西省", 
        "蒙": "内蒙古自治区", "辽": "辽宁省", "吉": "吉林省", "黑": "黑龙江省",
        "沪": "上海市", "苏": "江苏省", "浙": "浙江省", "皖": "安徽省",
        "闽": "福建省", "赣": "江西省", "鲁": "山东省", "豫": "河南省",
        "鄂": "湖北省", "湘": "湖南省", "粤": "广东省", "桂": "广西壮族自治区",
        "琼": "海南省", "渝": "重庆市", "川": "四川省", "黔": "贵州省",
        "滇": "云南省", "藏": "西藏自治区", "陕": "陕西省", "甘": "甘肃省",
        "青": "青海省", "宁": "宁夏回族自治区", "新": "新疆维吾尔自治区",
        "台": "台湾省", "港": "香港特别行政区", "澳": "澳门特别行政区"
    }
    
    for abbr, province in province_abbr.items():
        if abbr in location:
            return province
    
    # 特殊处理常见城市所属省份
    city_to_province = {
        "北京": "北京市", "上海": "上海市", "天津": "天津市", "重庆": "重庆市",
        "广州": "广东省", "深圳": "广东省", "杭州": "浙江省", "南京": "江苏省",
        "成都": "四川省", "武汉": "湖北省", "西安": "陕西省", "沈阳": "辽宁省"
    }
    for city, province in city_to_province.items():
        if city in location:
            return province
    
    return "未知省份"

def generate_province_colors(provinces):
    """为每个省份生成唯一的颜色"""
    # 生成鲜明且区分度高的颜色
    colors = []
    hue_steps = 360 / len(provinces) if provinces else 0
    
    for i, province in enumerate(provinces):
        # 使用HSV颜色空间，确保色调均匀分布，饱和度和明度保持较高值
        hue = int(i * hue_steps)
        # 转换HSV到RGB
        r, g, b = 0, 0, 0
        h = hue / 360.0
        s = 0.6  # 饱和度
        v = 0.8  # 明度
        
        i = int(h * 6)
        f = h * 6 - i
        p = v * (1 - s)
        q = v * (1 - f * s)
        t = v * (1 - (1 - f) * s)
        
        if i == 0:
            r, g, b = v, t, p
        elif i == 1:
            r, g, b = q, v, p
        elif i == 2:
            r, g, b = p, v, t
        elif i == 3:
            r, g, b = p, q, v
        elif i == 4:
            r, g, b = t, p, v
        else:
            r, g, b = v, p, q
        
        # 转换为十六进制
        r = int(r * 255)
        g = int(g * 255)
        b = int(b * 255)
        hex_color = f'#{r:02x}{g:02x}{b:02x}'
        colors.append((province, hex_color))
    
    return dict(colors)

def generate_dialect_json(excel_path, output_json="dialect_data.json"):
    """
    读取包含方言数据的Excel文件，按省份分配颜色并生成JSON格式数据
    
    参数:
    excel_path: Excel文件路径
    output_json: 输出的JSON文件路径
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(
            excel_path,
            usecols=["地点名", "纬度", "经度", "采集人", "音频", "内容"],
            header=0
        )
        
        # 提取省份信息
        df["省份"] = df["地点名"].apply(extract_province)
        
        # 获取所有唯一省份并生成颜色映射
        unique_provinces = df["省份"].unique().tolist()
        province_colors = generate_province_colors(unique_provinces)
        
        # 转换为列表字典格式
        dialect_list = []
        for index, row in df.iterrows():
            # 确保经纬度为数值类型
            try:
                longitude = float(row["经度"]) if pd.notnull(row["经度"]) else 0.0
                latitude = float(row["纬度"]) if pd.notnull(row["纬度"]) else 0.0
            except (ValueError, TypeError):
                print(f"警告: 第{index+1}行的经纬度格式不正确，已设为默认值")
                longitude = 0.0
                latitude = 0.0
            
            # 处理音频路径
            audio_url = row["音频"] if pd.notnull(row["音频"]) else ""
            audio_name = ""
            if audio_url:
                # 提取音频文件名（不含路径和扩展名）
                audio_name = os.path.splitext(os.path.basename(audio_url))[0]
            
            # 获取该省份对应的颜色
            province = row["省份"]
            color = province_colors.get(province, "#999999")  # 默认灰色
            
            # 创建数据对象，包含省份和颜色信息
            dialect_item = {
                "id": f"dialect_{index + 1}",
                "location": row["地点名"] if pd.notnull(row["地点名"]) else f"未知地点_{index + 1}",
                "province": province,  # 新增省份字段
                "longitude": longitude,
                "latitude": latitude,
                "color": color,  # 新增颜色字段（按省份分配）
                "collector": row["采集人"] if pd.notnull(row["采集人"]) else "未知采集人",
                "audio": {
                    "url": audio_url,
                    "name": audio_name
                },
                "content": row["内容"] if pd.notnull(row["内容"]) else "无内容描述"
            }
            
            dialect_list.append(dialect_item)
        
        # 保存为JSON文件
        with open(output_json, 'w', encoding='utf-8') as f:
            json.dump(dialect_list, f, ensure_ascii=False, indent=2)
        
        print(f"成功生成按省份着色的JSON数据: {output_json}")
        print(f"共包含 {len(dialect_list)} 条方言记录，涉及 {len(unique_provinces)} 个省份")
        return True
        
    except Exception as e:
        print(f"处理过程出错: {str(e)}")
        return False

if __name__ == "__main__":
    # 替换为你的Excel文件路径
    excel_file = "Data.xlsx"
    
    # 检查文件是否存在
    if not Path(excel_file).exists():
        print(f"错误: 找不到文件 {excel_file}")
    else:
        generate_dialect_json(excel_file)
