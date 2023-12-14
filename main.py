import requests
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl

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

def export_to_excel(danmaku_list, excel_file):
    df = pd.DataFrame({'弹幕': danmaku_list})
    df.to_excel(excel_file, index=False)

def main():
    # 用户输入视频ID
    video_id = input("请输入B站视频的ID: ")

    # 获取弹幕
    danmaku_xml = get_bilibili_danmaku(video_id)

    # 解析弹幕
    danmaku_list = parse_danmaku(danmaku_xml)

    # 导出到Excel
    export_to_excel(danmaku_list, "bilibili_danmaku.xlsx")

if __name__ == "__main__":
    main()
