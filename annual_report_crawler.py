import requests
import re
import pandas as pd
import csv
from urllib import request, parse
import json
import os

orgid_url = "http://www.cninfo.com.cn/new/information/topSearch/query"
# 文件名中过滤掉不需要下载的文件
exclude_file_arr = ['（已取消）', '年度报告摘要', '（更新前）']
flag = True  # 是否需要创建名称为新股票代码的文件夹
span = 8  # 时间跨度


def get_orgid(code, maxNum=10):
    request_param = {
        'keyWord': code,
        'maxNum': maxNum,
    }
    req_data = bytes(parse.urlencode(request_param, encoding='utf-8'), encoding='utf-8')
    req = request.Request(orgid_url, data=req_data, method="POST")
    res = request.urlopen(req)
    response = json.loads(res.read().decode('utf-8'))
    # 这是一个list，只拿出第1个即可
    if len(response) != 0:
        return response[0]['orgId']
    else:
        return None


def get_response(code, page_num, start_date, end_date):
    url = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'
    org_id = get_orgid(code)

    # if code[0] == '6':
    #     attachment = ',gssh0' + code
    # elif code[0] == '0':
    #     attachment = ',gssz0' + code
    # elif code[0] == '3':
    #     attachment = get_attachment(code)

    page_num = int(page_num)
    data = {
        'stock': code + ',' + org_id,
        'searchkey': '',
        'plate': '',
        'category': 'category_ndbg_szsh',
        'trade': '',
        'column': 'szse',
        'pageNum': page_num,
        'pageSize': 30,
        'tabName': 'fulltext',
        'sortName': '',
        'sortType': '',
        'limit': '',
        'showTitle': '',
        'seDate': start_date + '~' + end_date,
        'secid': ''
    }
    headers = {
        'Accept': '*/*',
        'Accept-Encoding': 'gzip,deflate',
        'Connection': 'keep-alive',
        'Content-Length': '181',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Host': 'www.cninfo.com.cn',
        'Origin': 'http://www.cninfo.com.cn',
        'Referer': 'http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }
    r = requests.post(url, data=data, headers=headers)
    result = r.json()['announcements']
    return result


def check_file_need_download(file_name):
    for exclude in exclude_file_arr:
        if exclude in file_name:
            return False
    return True


def filter_illegal_filename(file_name):
    illegal_char = {
        ' ': '',
        '*': '',
        '/': '-',
        '\\': '-',
        ':': '-',
        '?': '-',
        '"': '',
        '<': '',
        '>': '',
        '|': '',
        '－': '-',
        '—': '-',
        '（': '(',
        '）': ')',
        'Ａ': 'A',
        'Ｂ': 'B',
        'Ｈ': 'H',
        '，': ',',
        '。': '.',
        '：': '-',
        '！': '_',
        '？': '-',
        '“': '"',
        '”': '"',
        '‘': '',
        '’': ''
    }
    for item in illegal_char.items():
        file_name = file_name.replace(item[0], item[1])
    return file_name


def download_pdf(code, response):
    global flag
    if not response:
        print("股票代码为{0}的响应不存在，即不存在对应时间内的pdf文件".format(code))
    else:
        for item in response:
            need_download = check_file_need_download(item['announcementTitle'])  # 检查文件是否需要被下载
            if not need_download:
                print('###################过滤掉文件：' + item['announcementTitle'])
                continue
            else:
                # 自动生成文件夹，分门别类整理
                new_folder_name = item['secCode']  # 新文件夹命名为股票代码
                new_folder_path = "D:\\Enterprise financial crisis warning\\annual_report_download_pdf\\" + new_folder_name
                # 判断是否需要创建名称为新股票代码的文件夹
                if flag:
                    os.makedirs(new_folder_path)
                    # 检查新文件夹是否存在
                    if os.path.exists(new_folder_path):
                        print("命名为股票代码：{}的新文件夹已成功创建".format(new_folder_name))
                    else:
                        print("无法创建新文件夹")
                    flag = False

                # 确定pdf文件的命名格式
                title = item['announcementTitle']
                sec_name = item['secName'].replace('*', '')
                sec_code = item['secCode']
                file_name = f'{sec_code}_{sec_name}_{title}.pdf'
                file_name = filter_illegal_filename(file_name)
                save_path = new_folder_path + "\\" + file_name  # 确定pdf文件的保存路径
                # 确定pdf文件的下载链接
                adjunct_url = item['adjunctUrl']
                download_url = 'http://static.cninfo.com.cn/' + adjunct_url

                # 下载pdf文件
                r = requests.get(download_url)
                with open(save_path, 'wb') as f:
                    f.write(r.content)
                print(f'名为{sec_code}{sec_name}{title}的pdf文件已下载完毕')
        flag = True


if __name__ == '__main__':
    saving_path = 'D:\\Enterprise financial crisis warning\\'

    df = pd.read_excel('D:\\Enterprise financial crisis warning\\paired_result.xlsx',
                       sheet_name='paired_results', dtype={'Symbol': str, 'FirstSTDate': int})
    deduplicated_data = df.drop_duplicates(subset=['Symbol'], keep='first').reset_index(drop=True)
    code_list = deduplicated_data['Symbol'].tolist()

    for code in code_list:
        print('----------------------------------------------------------------------------')
        print('即将下载的是股票代码为{0}的年报'.format(code))
        # 获取起始，终止日期
        # 注意，这里end_date很难明确是多少（在这里是+5）。比如公司2003年的年报，将end_date设置到2007年才出现
        start_date = str(deduplicated_data[deduplicated_data['Symbol'] == code]['FirstSTDate'].iloc[0] - span) + '-01-01'
        end_date = str(deduplicated_data[deduplicated_data['Symbol'] == code]['FirstSTDate'].iloc[0] + 5) + '-01-01'
        print('下载年份为：{0} —— {1}'.format(start_date, end_date))
        # 若想爬取所有年份年报，修改 get_response 函数中的 start_date 和 end_date 即可
        # 如：response = get_response(code, 1, '2000-01-01', '2023-01-01')
        response = get_response(code, 1, start_date, end_date)
        download_pdf(code, response)



