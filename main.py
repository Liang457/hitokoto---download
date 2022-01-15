import sys
import requests
import xlwt
import json
import os
import time


def main_get(frequency):
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet("Main")

    worksheet.write(0, 0, '序列')
    worksheet.write(0, 1, '一言标识')
    worksheet.write(0, 2, '一言正文')
    worksheet.write(0, 3, '类型')
    worksheet.write(0, 4, '一言的出处')

    for i in range(1, frequency + 1):
        text = get_one()
        json_text = json_analysis(text)

        worksheet.write(i, 0, str(i))  # 写入序列
        worksheet.write(i, 1, json_text['id'])  # 写入一言标识
        worksheet.write(i, 2, json_text['hitokoto'])  # 写入一言正文
        worksheet.write(i, 3, json_text['type'])  # 写入类型
        worksheet.write(i, 4, json_text['from'])  # 写入一言的出处

        progress_bar(frequency, i, json_text['hitokoto'])
        workbook.save("一言.xls")
        time.sleep(0.1)

    os.startfile('一言.xls')


def get_one():
    url = 'https://international.v1.hitokoto.cn'

    word = requests.get(url)
    word.encoding = 'utf-8'
    return word.text


def json_analysis(text):
    return json.loads(text)


def progress_bar(all_t, now_t, now_w):
    os.system('cls')
    bai_fen_bi = now_t / all_t
    bai_fen_bi = "%.2f%%" % (bai_fen_bi * 100)
    print(f'''爬取完成
一言正文:{now_w}
爬取进度 {now_t}/{all_t}[{bai_fen_bi}]
''')


main_get(int(input('爬取多少次？')))
