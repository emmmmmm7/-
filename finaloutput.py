import pandas as pd
import re
import json
import copy

# DMS 转换为十进制度的函数
def dms_to_decimal(dms_str):
    """将DMS格式(101/45/10.00)转换为十进制度"""
    match = re.match(r'(\d+)/(\d+)/([\d.]+)', str(dms_str))
    if match:
        degrees = int(match.group(1))
        minutes = int(match.group(2))
        seconds = float(match.group(3))
        return round(degrees + minutes / 60 + seconds / 3600, 6)
    return None

# 加载 Excel 文件
file_path = 'yourfile.xlsx'  # 更新为你的文件路径
df = pd.read_excel(file_path, header=None, engine='openpyxl')  # 读取文件，无表头

# 初始化变量
last_coordinates = {"Longitude": None, "Latitude": None}
last_sample_id = None
last_row_id = 0

# 遍历数据并合并未标明坐标的样品号
for i, row in df.iterrows():
    sample_id = row[0]  # 第一列：样品号
    longitude = row[1]  # 第二列：经度
    latitude = row[2]   # 第三列：纬度

    if pd.isnull(longitude) or pd.isnull(latitude):
        # 如果当前行没有坐标，合并样品号
        df.at[i, 1] = last_coordinates['Longitude']
        df.at[i, 2] = last_coordinates['Latitude']
        last_sample_id = f"{last_sample_id.split('-')[0]}-{last_sample_id.split('-')[1]},{sample_id.split('-')[1]}"
        df.at[last_row_id, 0] = last_sample_id      
    else:
        # 更新最新的坐标和样品号
        last_coordinates["Longitude"] = dms_to_decimal(longitude)
        last_coordinates["Latitude"] = dms_to_decimal(latitude)
        last_sample_id = sample_id
        last_row_id = i

# 将坐标转换为十进制度
df[1] = df[1].apply(dms_to_decimal)  # 第二列转换为十进制度
df[2] = df[2].apply(dms_to_decimal)  # 第三列转换为十进制度

# 删除 Longitude 列中为空值的行
df_cleaned = df[df.iloc[:, 1].notna()]
# 保存结果到新文件
output_path = 'yourfile.csv'
df_cleaned.to_csv(output_path, index=False, header=False)
print(f"处理后的数据已保存到 {output_path}")

# 将清理后的数据转换为代码1中的输入格式
data = "\n".join(
    f"{row[0]}\t{row[1]}\t{row[2]}" for row in df_cleaned.itertuples(index=False)
)

# 基础模板结构
base_template = {
    "Version": "V9.3.0",
    "Type": 1,
    "ObjItems": [
        {
            "Type": 30,
            "ObjID": 880536910,
            "ParentID": 1361484499,
            "SrvID": "4723318442077404832",
            "tmModify": "2024/01/17 15:43:31",
            "Object": {
                "Name": "2020-A-型",
                "Type": 30,
                "ObjectDetail": {
                    "Child": 26,
                    "LoadOk": 1,
                    "SaveMerge": 0,
                    "Group": 0,
                    "AutoLoad": 1,
                    "ShowLevel": 1,
                    "ShowLevelMax": 0,
                    "Crypt": 0,
                    "Share": 0,
                    "ReadOnly": 0,
                    "NotHotId": 0,
                    "Bind": 0,
                    "BindCheck": 0,
                    "Link": 0,
                    "LinkAutoCheck": 0,
                    "ChildiOverlay": 0,
                    "LinkUrl": "",
                    "Relate": 0,
                    "ObjChildren": []
                }
            }
        }
    ]
}

# 子对象模板
child_template = {
    "Type": 7,
    "SrvID": "215308488163548131",  # 示例ID，实际可生成随机唯一值
    "ObjID": 813801718,  # 示例ID，实际可生成随机唯一值
    "tmModify": "2021/11/30 20:54:43",
    "ParentID": 880536910,
    "Object": {
        "Name": "",
        "Type": 7,
        "Comment": "",
        "ObjectDetail": {
            "Lat": 0,
            "Lng": 0,
            "Gcj02": 1,
            "Altitude": 0,
            "EditMode": 0,
            "OverlayIdx": 0,
            "TxtType": 1,
            "ShowLevel": 1,
            "ShowLevelMax": 0,
            "TimeUncertain": 0,
            "SignEvent": {"Radius": 0, "ShowClr": 0},
            "SignPic": {
                "SignPic": 10,
                "AlignFlag": 0,
                "SignClr": 0,
                "PicScale": 10,
                "SignPicNum": 0,
                "SignPicNumOffx": 0,
                "SignPicNumOffy": 0,
                "SignPicNumClr": 0,
                "SignPicNumSize": 0,
            },
            "TxtShowSta": 0,
            "TxtShowStaSet": 0,
        },
    },
}

# 数据处理：填充子对象
obj_children = []
for idx, line in enumerate(data.strip().split("\n")):
    parts = line.split("\t")
    name, lat, lng = parts[0], float(parts[2]), float(parts[1])

    # 复制子对象模板并填充数据
    obj = copy.deepcopy(child_template)
    obj["SrvID"] = f"2153084881635481{idx:02d}"  # 示例唯一ID
    obj["ObjID"] = 813801718 + idx  # 示例唯一ID
    obj["Object"]["Name"] = name
    obj["Object"]["ObjectDetail"]["Lat"] = lat
    obj["Object"]["ObjectDetail"]["Lng"] = lng

    obj_children.append(obj)

# 将子对象添加到基础模板中
base_template["ObjItems"][0]["Object"]["ObjectDetail"]["ObjChildren"] = obj_children

# 输出到 .ovjson 文件
output_file_path = "yourfile.ovjsn"
with open(output_file_path, "w", encoding="utf-8") as file:
    json.dump(base_template, file, indent=4, ensure_ascii=False)

print(f"数据已成功保存到文件：{output_file_path}")
