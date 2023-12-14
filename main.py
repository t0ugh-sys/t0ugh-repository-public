import requests
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import messagebox

def get_bilibili_danmaku(video_id):
    danmaku_url = f"https://api.bilibili.com/x/v1/dm/list.so?oid={video_id}"

    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        response = requests.get(danmaku_url, headers=headers)
        response.raise_for_status()
        return response.content
    except requests.exceptions.HTTPError as errh:
        print("HTTP Error:", errh)
    except requests.exceptions.ConnectionError as errc:
        print("Error Connecting:", errc)
    except requests.exceptions.Timeout as errt:
        print("Timeout Error:", errt)
    except requests.exceptions.RequestException as err:
        print("Something went wrong:", err)

def parse_danmaku(xml_content):
    danmaku_list = []
    root = ET.fromstring(xml_content)
    for d in root.findall(".//d"):
        danmaku_list.append(d.text)
    return danmaku_list

def export_to_excel(danmaku_list, video_id):
    excel_file = f"bilibili_danmaku_{video_id}.xlsx"
    df = pd.DataFrame({'弹幕': danmaku_list})
    df.to_excel(excel_file, index=False)
    messagebox.showinfo("导出成功", f"弹幕已成功导出到 {excel_file}")

def get_video_id():
    video_id = entry.get()
    return video_id

def on_export_button_click():
    video_id = get_video_id()
    danmaku_xml = get_bilibili_danmaku(video_id)
    danmaku_list = parse_danmaku(danmaku_xml)
    export_to_excel(danmaku_list, video_id)

# 创建主窗口
root = tk.Tk()
root.title("Bilibili弹幕导出")

# 设置窗口大小
root.geometry("300x200")

# 添加标签和输入框
label = tk.Label(root, text="B站视频ID:")
label.pack(pady=10)
entry = tk.Entry(root)
entry.pack(pady=10)

# 添加导出按钮
export_button = tk.Button(root, text="导出弹幕到Excel", command=on_export_button_click)
export_button.pack(pady=20)

# 运行主循环
root.mainloop()
