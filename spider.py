import math
import re
from datetime import datetime
import requests
import pandas as pd
import json
import time  # 用于设置爬取间隔


def crawl_baoyan_data(url):
    try:
        # 发送请求获取数据
        response = requests.get(url)
        response.raise_for_status()  # 检查请求是否成功
        data = response.json()

        # 提取需要的内容列表和总页数
        result = {}
        if data.get('success') and 'result' in data:
            result['content'] = data['result'].get('content', [])
            # result['total_pages'] = data['result'].get('total_pages', 0)  # 获取总页数
            result['page_size'] = data['result']['page_size']
            result['total_count'] = data['result']['total_count']
            #     page_size = first_page_result['page_size']
            #     total_count = first_page_result['total_count']

            return result
        else:
            print("获取数据失败，返回格式异常")
            return {'content': [], 'total_pages': 0}
    except Exception as e:
        print(f"爬取过程出错：{e}")
        return {'content': [], 'total_pages': 0}


def process_data(content_list):
    processed_data = []
    for item in content_list:
        # 处理content字段中的嵌套JSON
        content_str = item.get('content', '{}')
        try:
            content_json = json.loads(content_str)
            # 提取content中的文本内容（p字段）
            content_text = content_json.get('p', '').replace('\n', ' ').strip()
        except json.JSONDecodeError:
            content_text = content_str.replace('\n', ' ').strip()
        bracket_pattern = re.compile(r'【([^】]+)】')
        # 整理需要保存的字段（可根据需求增删）
        processed_item = {
            'id': item.get('id', ''),
            '标题': item.get('title', ''),
            '发布时间': item.get('created_at', ''),
            '学校': (
                item.get('college', '').strip()
                or next(iter(bracket_pattern.findall(item.get('title', ''))), '')
            ),
            '学院': item.get('academy', ''),
            '专业': item.get('major', ''),
            '省份': item.get('province', ''),
            '学校层次': item.get('college_level', ''),
            '标签': item.get('tags', ''),
            '报名开始时间': item.get('sign_up_start', ''),
            '报名结束时间': item.get('sign_up_end', ''),
            '活动开始时间': item.get('start_time', ''),
            '活动结束时间': item.get('end_time', ''),
            '官方链接': item.get('office_url', ''),
            '报名链接': item.get('sign_up_url', ''),
            '内容摘要': content_text,
            '申请人数': item.get('apply_cnt', ''),
            '浏览人数': item.get('watch_cnt', '')
        }
        processed_data.append(processed_item)
    return processed_data


def save_to_excel(data, filename='保研信息数据.xlsx'):
    if not data:
        print("没有数据可保存")
        return
    df = pd.DataFrame(data)
    # 保存为Excel
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"数据已成功保存到 {filename}，共 {len(data)} 条记录")


def crawl_all_pages(base_url, delay_seconds=3):
    """
    爬取所有页面的数据
    :param base_url: 基础URL（包含page=占位符）
    :param delay_seconds: 每页爬取后的延迟时间（秒）
    :return: 所有页面的合并数据
    """
    all_content = []
    page = 1

    # 先获取第一页数据，确定总页数
    first_page_url = base_url.format(page)
    print(f"开始爬取第1页：{first_page_url}")
    first_page_result = crawl_baoyan_data(first_page_url)
    all_content.extend(first_page_result['content'])
    page_size = first_page_result['page_size']
    total_count = first_page_result['total_count']
    total_pages = math.ceil(total_count/page_size)

    if total_pages <= 1:
        print("仅发现1页数据，无需继续爬取")
        return all_content

    print(f"共发现{total_pages}页数据，开始爬取剩余页面...")

    # 爬取剩余页面
    for page in range(2, total_pages + 1):
        # 构造当前页URL
        current_url = base_url.format(page)
        print(f"开始爬取第{page}页：{current_url}")

        # 爬取当前页数据
        page_result = crawl_baoyan_data(current_url)
        all_content.extend(page_result['content'])

        # 页面爬取间隔（最后一页不需要延迟）
        if page < total_pages:
            print(f"第{page}页爬取完成，等待{delay_seconds}秒后继续...")
            time.sleep(delay_seconds)

    print(f"所有{total_pages}页数据爬取完成，共获取{len(all_content)}条记录")
    return all_content


if __name__ == "__main__":
    # 目标API地址（注意修改为带page占位符的格式）
    base_url = "http://api.baoyanwang.com.cn/api/v1/articles?page={}&size=25&category=%E4%BF%9D%E7%A0%94%E4%BF%A1%E6%81%AF&all=1"
    # http://api.baoyanwang.com.cn/api/v1/articles?page=1&size=25&category=%E4%BF%9D%E7%A0%94%E4%BF%A1%E6%81%AF&all=1
    # 爬取所有页面数据（每页间隔3秒）
    all_content = crawl_all_pages(base_url, delay_seconds=3)

    # 处理数据
    processed_data = process_data(all_content)

    # 保存为Excel

    today = datetime.now().strftime('%Y%m%d')
    filename = f'保研信息全量数据_{today}.xlsx'
    save_to_excel(processed_data, filename)