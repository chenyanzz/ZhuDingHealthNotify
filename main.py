# -*- coding: UTF-8 -*-

# 请在钉钉群里添加钉钉机器人，由于安全考虑，需要设置关键词，可以设置为每条消息都有的 "健康打卡"几个字

import datetime
import json
import math
import os
import shutil
import sys
import logging
import time
from collections import defaultdict
from http.cookiejar import MozillaCookieJar

import requests
import xlrd

import zju_login

# 请将账号密码写入 config.json
config = {}

def readCfg():
    with open("config.json", "rt") as f:
        global config
        config = json.loads(f.read())

map_misstime = {}
map_mobile = {}


def fill_list(my_list: list, length, fill=None):  # 使用 fill字符/数字 填充，使得最后的长度为 length
    if len(my_list) >= length:
        return my_list
    else:
        return my_list + (length - len(my_list)) * [fill]


def read_misstime():
    global config
    global map_misstime
    with open("misstime.list", "r", encoding="UTF-8") as f:
        lines = f.readlines()
        for line in lines:
            v = line.split(" ")
            if len(v)<=3:
                continue
            name = v[0]
            mobile = v[1]
            _miss_arr = fill_list(v[2:], int(config["misstime_dayrange"]), 0)
            miss_arr = [int(v) for v in _miss_arr]
            map_misstime[name] =  miss_arr[1:] + [0]
            map_mobile[name] = mobile

def save_misstime():
    global config
    global map_misstime

    with open("misstime.list", "w+", encoding="UTF-8") as f:
        for name, misstime_arr in map_misstime.items():
            v = [name, map_mobile[name]] + [str(v) for v in misstime_arr]
            line = " ".join(v)+"\n"
            f.write(line)
    print(map_misstime)
    print(map_mobile)

# 简单的下载函数，未做retry


def download_file(sess, url, out_file):
    logging.info('输出到文件： %s' % out_file)

    response = sess.get(url, stream=True, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"
    })
    with open(out_file, 'wb') as f:
        response.raw.decode_content = True
        shutil.copyfileobj(response.raw, f)
    return out_file


def get_date_str():
    return datetime.datetime.now().strftime('%Y%m%d')


def get_datetime_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')


def sub_send_msg(message, at_mobiles, ding_robot_url, label):
    post_json = {
        "msgtype": "text",
        "text": {
            "content": message,
        },
        "at": {
            "atMobiles": at_mobiles,
            "isAtAll": False
        }
    }

    r = requests.post(ding_robot_url, json=post_json)

    logging.info(label + r.json()['errmsg'])

    return r.json()


# 批量提醒
def send_normal_ding_msg(person_list, ding_robot_url):
    if len(person_list) == 0:
        return

    results = []
    # print(person_list, ding_robot_url)

    # 提取需要@的手机号
    at_mobiles = [p["mobile"] for p in person_list]

    chunk_size = 50
    start = 0
    chunk_num = math.ceil(len(at_mobiles) / chunk_size)
    current_idx = 1

    while start < len(at_mobiles):
        end = min(len(at_mobiles), start + chunk_size)
        loc_mobile = at_mobiles[start:end]
        message = f"以下同学请尽快健康打卡（消息 {current_idx}/{chunk_num})：\n打卡链接：https://healthreport.zju.edu.cn/ncov/wap/default/index\n"
        label = datetime.datetime.now().strftime('%Y-%m-%d') + \
            "，第({} / {})条钉钉消息发送结果".format(current_idx, chunk_num)
        res = sub_send_msg(message, loc_mobile, ding_robot_url, label)
        results.append(res)
        start += chunk_size
        current_idx += 1

    for result in results:
        if result["errcode"] != 0:
            logging.error(result["errmsg"])

    return results

# 给那些常常不打卡的

def send_VIP_ding_msg(person_list, ding_robot_url):
    message = "据说，以下同学常常忘记健康打卡：\n\n"
    at_mobiles = []

    global map_misstime
    global config

    logging.info("开始VIP钉")

    # 方法1: 提醒当天屡教不改的
    # for p in person_list:
    #     name = p["name"]
    #     mobile = p["mobile"]

    # 方法2: 提醒一周屡教不改的
    for name in map_misstime.keys():
        mobile = map_mobile[name]
        mt = int(config["misstime_ding_mintime"])
        cnt = sum(map_misstime[name]) if (name in map_misstime) else 0
        if  cnt >= mt:
            logging.info(f"VIP提醒:{name+mobile}")
            message += f"{name}\t: 在{ int(config['misstime_dayrange']) }天里忘打卡{cnt}次啦！\n"
            at_mobiles.append(mobile)

    if len(at_mobiles) == 0:
        logging.info("没有VIP同学")
        return
    sub_send_msg(message, at_mobiles, ding_robot_url, "特别提醒")

    logging.info("VIP钉成功")


def set_logging():
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    while root_logger.hasHandlers():
        for i in root_logger.handlers:
            root_logger.removeHandler(i)
    root_logger.addHandler(
        logging.FileHandler(filename='logs/' + datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S') + '.log',
                            encoding='utf-8'))
    logging.basicConfig(format='[%(asctime)s-%(levelname)s:%(message)s]')


def downloadXlsxFile():
    logging.info(
        "正在下载" + datetime.datetime.now().strftime('%Y%m%d') + "疫情未上报情况文件")

    # 持久化cookies
    sess = requests.Session()
    cookie_file = "cookies.txt"
    cookies = MozillaCookieJar(cookie_file)
    if os.path.exists(cookie_file):
        cookies.load(ignore_discard=True, ignore_expires=True)
        sess.cookies = cookies

    # 未上报报表下载url
    url = "https://healthreport.zju.edu.cn/ncov/wap/zju/export-download?group_id=1&group_type=1&type=weishangbao&date={}".format(
        get_date_str())
    dest_file = os.path.join("records", "{}.xlsx".format(get_datetime_str()))
    # 下载文件
    zju_login.login(
        sess, username=config["username"], password=config["password"])
    logging.info("登录成功")
    record_file = download_file(sess, url, dest_file)
    logging.info("文件下载完成：" + str(record_file))

    return record_file


def getStuData(record_file):
    wb = xlrd.open_workbook(record_file)
    sheet = wb.sheet_by_index(0)
    # 构造表头
    headers = dict((i, sheet.cell_value(0, i)) for i in range(sheet.ncols))

    # 提取信息
    values = list(
        (
            dict((headers[j], sheet.cell_value(i, j))
                 for j in headers)
            for i in range(1, sheet.nrows)
        )
    )

    return values


def addMissTime(name, is_miss=True):
    global map_misstime
    l = int(config["misstime_dayrange"])
    if name not in map_misstime:
        miss_arr = fill_list([], l, 0)
        map_misstime[name] = miss_arr

    map_misstime[name][l-1] = 1

def refreshMissTime(ding_list):
    for stu in ding_list:
        name = stu["name"]
        mobile = stu["mobile"]
        addMissTime(name)
        map_mobile[name] = mobile

def stu_data2ding_list(stu_data):
    exclude_sid_list = []
    with open("excludes.txt", "rt") as f:
        exclude_sid_list = f.read().strip().splitlines()

    group_by_grade = defaultdict(list)

    for row in stu_data:
        name = row["姓名"]
        sid = row["学工号"]
        mobile = row["手机号码"]
        grade = sid[1:3]  # 从学号中提取年级（只适配了研究生的学号，如 22051001 20为年级）

        if sid in exclude_sid_list:
            logging.warning("跳过：" + name + sid + "，(在排除列表中)")
            continue
        if not mobile:
            logging.warning("跳过：" + name + sid + "，没有手机号信息")
            continue
        # 先把需要提醒的收集起来
        group_by_grade["21"].append({"name": name, "mobile": mobile})

    return group_by_grade

if __name__ == '__main__':
    readCfg()
    set_logging()

    logging.info("开始打卡提醒")

    filename = downloadXlsxFile()
    stu_data = getStuData(filename)

    ding_list  = stu_data2ding_list(stu_data)["21"]

    logging.info(ding_list)

    if sys.argv[1] == "day":
        logging.info("通知21级学生")
        robot_url = config["grade_group_robot_mapping"]["21"]
        send_normal_ding_msg(ding_list, robot_url)
        read_misstime()
        send_VIP_ding_msg(ding_list, robot_url)

    elif sys.argv[1] == "night":
        logging.info("更新未打卡名单")
        if datetime.date.today().weekday!=6:
            read_misstime()
        refreshMissTime(ding_list)
        save_misstime()
        