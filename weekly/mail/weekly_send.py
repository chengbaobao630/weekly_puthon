#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
from datetime import timedelta, datetime
from enum import Enum
import pythoncom

import mammoth
import win32com.client as win32
from bs4 import BeautifulSoup
from docx import Document


class DateType(Enum):
    this_week = "this_week"
    next_week = "next_week"
    after = "after"


now = datetime.now()
this_week_start = now - timedelta(days=now.weekday())
this_week_end = now + timedelta(days=6 - now.weekday())

weeklyData = ""


def init(data):
    global weeklyData
    weeklyData = data


def get_job(type_in, num) -> str:
    try:
        return weeklyData[type_in.name]["job"][int(num)]["job"]
    except Exception as e:
        print(e)
        return ""


def get_content(type_in, num) -> str:
    try:
        return weeklyData[type_in.name]["job"][int(num)]["content"]
    except Exception as e:
        print(e)
        return ""


def get_finish(type_in, num) -> str:
    try:
        return "{0}%".format(weeklyData[type_in.name]["job"][int(num)]["percent"])
    except Exception as e:
        print(e)
        return ""


def get_unfinished(type_in, num) -> str:
    try:
        unfinished = 100 - int(weeklyData[type_in.name]["job"][int(num)]["percent"])
        if unfinished <= 0: return "无"
        return unfinished + "%"
    except Exception as e:
        print(e)
        return ""


def get_time(type_in, num) -> str:
    try:
        return weeklyData[type_in.name]["job"][int(num)]["time"]
    except Exception as e:
        print(e)
        return ""


def get_result(type_in, num) -> str:
    try:
        return weeklyData[type_in.name]["job"][int(num)]["result"]
    except Exception as e:
        print(e)
        return ""


def get_unfinished_trace() -> str:
    try:
        return weeklyData[DateType.this_week.name]["trace"]
    except Exception as e:
        print(e)
        return ""


def get_qa() -> str:
    try:
        return weeklyData[DateType.this_week.name]["qa"]
    except Exception as e:
        print(e)
        return ""


def get_next_week_expect_result() -> str:
    try:
        return weeklyData[DateType.next_week.name]["expect_result"]
    except Exception as e:
        print(e)
        return ""


def get_support() -> str:
    try:
        return weeklyData[DateType.next_week.name]["support"]
    except Exception as e:
        print(e)
        return ""


this_week = {
    "submit_time": this_week_end.strftime("%y/%m/%d"),
    "start_time_end_time": "{0}至{1}".format(this_week_start.strftime("%y/%m/%d"),
                                            this_week_end.strftime("%y/%m/%d")),
    "job_": get_job,
    "content_": get_content,
    "finish_": get_finish,
    "unfinished_": get_unfinished,
    "unfinished_trace": get_unfinished_trace,
    "Q&A": get_qa,
    "support": get_support,
    "result_": get_result,
}

next_week = {
    "job_": get_job,
    "job_content_": get_content,
    "job_time_": get_time,
    "expect_result": get_next_week_expect_result
}

after = {
    "job_": get_job,
    "job_content_": get_content,
    "job_time_": get_time
}


def do(cell_param, type_in, week_params, prefix):
    for key, value in week_params.items():
        print("key:{0},value:{1}".format(key, value))
        if prefix:
            key = prefix + key
        if key.endswith("_"):
            key += "."
        m_in = re.compile(".*{0}.*".format("{" + key + "}"))
        if re.search(m_in, cell_param.text):
            if type(value) == str:
                cell_param.text = cell_param.text.replace("{" + key + "}", value)
            else:
                # cell_param.text = value()
                rs = re.search("\d", cell_param.text)
                if rs and rs.group():
                    real_value = value(type_in, rs.group())
                    if type(real_value) == int:
                        real_value = str(real_value)
                    cell_param.text = real_value
                else:
                    cell_param.text = value()


def do_this_week(cell_param):
    do(cell_param, DateType.this_week.this_week, this_week, "")


def do_next_week(cell_param):
    do(cell_param, DateType.next_week, next_week, "next_week_")


def do_after(cell_param):
    do(cell_param, DateType.after, after, "after_")


def handler(cell_param):
    do_this_week(cell_param)
    do_next_week(cell_param)
    do_after(cell_param)


def sendemail(sub, body, file_path):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    receivers = ['林荣波(觉醒) <linrb@yoozoo.com>']
    mail = outlook.CreateItem(0x0)
    mail.To = receivers[0]
    mail.CC = '林帆(林青侠) <linfan@yoozoo.com>; 刘思远(刘工) <liusiyuan@yoozoo.com>;' \
              ' 丁丽盈(丁丁) <dingly@yoozoo.com>; 兰旭(布鲁斯) <lanx@yoozoo.com>'
    # mail.BCC = ['1060471903@qq.com', "799096114@qq.com"]
    mail.Subject = sub
    # mail.Body = MIMEText(body)
    mail.HTMLBody = body
    # 添加附件
    mail.Attachments.Add(file_path)
    mail.Send()


def gen_send(data):
    init(data)
    subject = "{0}_工作周报_{1}~{2}".format("程呈", this_week_start.strftime("%y-%m-%d"),
                                        this_week_end.strftime("%y-%m-%d"))
    file_name = "{0}.doc".format(subject)
    document = Document(r"C:\Users\chcheng\PycharmProjects\weekly\weekly\mail\weeklyTemplate.docx")
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                handler(cell)
    document.save(file_name)
    # time.sleep(2)
    with open(file_name, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # The generated HTML
        messages = result.messages  # Any messages, such as warnings during conversion
        soup = BeautifulSoup(html, 'html.parser')
        table_html = soup.table
        table_html["border"] = "1px solid"
        table_html["style"] = "border-collapse: collapse"
        for tr in soup.find_all("tr"):
            tds = tr.contents
            td_length = len(tds)
            if td_length < 5:
                if td_length >= 2:
                    tds[1]["colspan"] = 6 - td_length
                elif td_length == 1:
                    tds[0]["colspan"] = 5
                else:
                    continue

        sendemail(subject, """<html>
                                    <head>
                                    </head>
                                    <body>
                                        {0}
                                    </body>
                                </html>""".format(soup.prettify()),
                  os.path.dirname(os.path.realpath(file_name)) + os.path.altsep + file_name)
