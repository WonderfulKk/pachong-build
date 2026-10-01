import requests
import pandas as pd
import time
import os
import sys

# ------------------ 第一步：交互输入 ------------------
print("请提供以下信息（输入后按回车）：")

LIST_API_URL = input("列表页URL地址: ").strip()
LIST_USER_AGENT = input("列表页User-Agent: ").strip()
LIST_AUTH_TOKEN = input("列表页Authorization Token (例如 Bearer xxx): ").strip()

DETAIL_API_URL = input("详情页URL地址: ").strip()
DETAIL_USER_AGENT = input("详情页User-Agent: ").strip()
DETAIL_AUTH_TOKEN = input("详情页Authorization Token: ").strip()

print("\n参数已接收，开始爬取数据...\n")

# ------------------ 第二步：请求列表页 ------------------
headers_list = {
    "User-Agent": LIST_USER_AGENT,
    "Authorization": LIST_AUTH_TOKEN
}

try:
    resp_list = requests.get(LIST_API_URL, headers=headers_list, timeout=15)
    resp_list.raise_for_status()
    list_data = resp_list.json()
except Exception as e:
    print(f"列表页请求失败: {e}")
    input("按回车键退出...")
    sys.exit(1)

# ------------------ 第三步：解析列表页，提取信号列表 ------------------
if list_data.get('code') != 0:
    print(f"列表页返回错误: {list_data.get('msg')}")
    input("按回车键退出...")
    sys.exit(1)

records = list_data.get('data', {}).get('records', [])
if not records:
    print("列表页 records 为空，无数据可爬取。")
    input("按回车键退出...")
    sys.exit(0)

print(f"共获取到 {len(records)} 条信号，开始爬取处置记录...")

# ------------------ 第四步：循环请求详情页，并组装 Excel 行 ------------------
headers_detail = {
    "User-Agent": DETAIL_USER_AGENT,
    "Authorization": DETAIL_AUTH_TOKEN
}

excel_rows = []  # 用于存放最终要写入 Excel 的每一行数据

for idx, signal in enumerate(records):
    biz_id = signal.get('bizId')  # 列表页中用于请求详情页的标识
    if not biz_id:
        print(f"第 {idx+1} 条信号缺少 bizId，跳过")
        continue

    # 提取列表页中需要的字段（根据实际字段名调整）
    station_name = signal.get('stationName', '')
    signal_time = signal.get('signalOccurTime', '')
    task_content = signal.get('taskContent', '')

    print(f"正在处理第 {idx+1}/{len(records)} 条信号，bizId: {biz_id}")

    payload = {'mainId': biz_id}  # 详情页需要 mainId，其值为列表页的 bizId

    disposal_combined = ""  # 用于存放合并后的处置记录文本

    try:
        resp_detail = requests.post(
            DETAIL_API_URL,
            json=payload,          # 如果接口要求表单格式，改为 data=payload
            headers=headers_detail,
            timeout=10
        )
        resp_detail.raise_for_status()
        detail_data = resp_detail.json()

        # 提取处置记录（根据实际结构调整路径）
        # 假设处置记录在 detail_data['data']['data'] 中
        disposal_list = detail_data.get('data', {}).get('data', [])

        if disposal_list:
            disposal_texts = []
            for disposal in disposal_list:
                # 按指定格式拼接每条处置记录，每个字段之间用两个空格隔开
                text = (
                    f"【{disposal.get('process', '')}】  "
                    f"时间：{disposal.get('processTime', '')}  "
                    f"部门：{disposal.get('processDepartmentName', '')}  "
                    f"处理人：{disposal.get('processUserName', '')}  "
                    f"详情：{disposal.get('processDesc', '')}"
                )
                disposal_texts.append(text)
            # 每条记录换行，合并到一个字符串
            disposal_combined = "\n".join(disposal_texts)
            print(f"    -> 获取到 {len(disposal_list)} 条处置记录")
        else:
            print(f"    -> 该信号暂无处置记录")
            disposal_combined = "无"

    except Exception as e:
        print(f"    -> 请求或解析失败: {e}")
        disposal_combined = "请求失败"

    # 将本条信号的数据添加到 Excel 行列表中
    excel_rows.append({
        "变电站": station_name,
        "信号发出时间": signal_time,
        "内容": task_content,
        "处置记录": disposal_combined
    })

    time.sleep(1)  # 礼貌性暂停

# ------------------ 第五步：保存到 Excel ------------------
if excel_rows:
    df = pd.DataFrame(excel_rows, columns=["变电站", "信号发出时间", "内容", "处置记录"])
    # 保存到桌面
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop, "信号处置记录.xlsx")
    df.to_excel(file_path, index=False, engine='openpyxl')
    print(f"\n成功！共导出 {len(df)} 条信号记录，文件位于：{file_path}")
else:
    print("\n未获取到任何数据，请检查接口或数据结构。")

input("\n按回车键退出...")  # 防止运行完后窗口立即关闭