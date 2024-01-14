# encoding: utf-8
import common
import crawl
import pandas as pd
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import tqdm
import logging
import os
import time

################ Set up parameters here #####################
# 浏览器默认下载路径设置
# The first string should be the path where your file is downloaded to by default.
# Most likely, it should be like: "C://Users/usr_nm/Downloads"
default_download_path = "C:/Users/yunruxian/Downloads" + "/savedrecs.txt"

# 搜索任务所在excel文件路径设置
task_path = "raw data file/teachers‘list(20240106).xls"

# 读取搜索任务文件
df = pd.read_excel(
    task_path,
    sheet_name="Sheet1",
    header=0,
    keep_default_na=False,
)

# 获得搜索任务列表
task_list = []
for i in range(0, len(df)):
    school_id = df.iloc[i]["学校ID"]
    school = df.iloc[i]["学校"]
    teacher_id = df.iloc[i]["教师ID"]
    teacher_name_cn = df.iloc[i]["姓名"]
    teacher_name_en = df.iloc[i]["name"]
    address_en = df.iloc[i]["address"]
    address_en_plus = df.iloc[i]["address_plus"]
    others_information = {"学校ID": school_id, "教师ID": teacher_id, "姓名": teacher_name_cn}
    task_list.append([f"{school}", F"AU='{teacher_name_en}' AND AD='{address_en}'", others_information])
    if address_en_plus:
        task_list.append([f"{school}", F"AU='{teacher_name_en}' AND AD='{address_en_plus}'", others_information])

"""
Start the search of all tasks.
driver: the handle of a selenium.webdriver object
task_list: the zip of save paths and advanced query strings
default_download_path: the default path set for the system, for example, C://Downloads/
"""

# Init
"""设置logging"""
os.makedirs('logs', exist_ok=True)
logging.basicConfig(level=logging.INFO,
                    filename=os.getcwd() + '/logs/log' + time.strftime('%Y%m%d%H%M',
                                                                       time.localtime(time.time())) + '.log',
                    filemode="w",
                    format="%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s"
                    )

if __name__ == '__main__':
    threader = ThreadPoolExecutor(max_workers=5)
    remaining_tasks = []
    # Start Query
    # tqdm.tqdm(task_list)生成一个由迭代对象组成的进度条
    for school_folder, query_task, others_information in tqdm.tqdm(task_list):
        # 传统爬取命令
        # crawl.start(school_folder, query_task, default_download_path, others_information)
        # 多线程命令
        future = threader.submit(crawl.start, school_folder, query_task, default_download_path, others_information)
        remaining_tasks.append(future)

    # 线程计数
    while True:
        for future in remaining_tasks:
            if future.done():
                remaining_tasks.remove(future)
        common.show_progress(4118, (4118 - len(remaining_tasks)))
        # 所有任务完成时就退出循环
        if len(remaining_tasks) == 0:
            break

    # 等待所有的线程对象完成
    # wait的参数就是一个列表，运行到这里时，列表里已经没有对象了，所以下面这条命令重复
    # concurrent.futures.wait(remaining_tasks)
    threader.shutdown()

    # 检查是否全部完成
    for school_folder, query_task, others_information in tqdm.tqdm(task_list):
        path_school = os.path.join("downloads", school_folder)
        if not common.Check.mark_task_finish_flag(path_school, query_task):
            print(school_folder, query_task)

    # trans form json to excel
    for school_folder, query_task, others_information in tqdm.tqdm(task_list):
        if os.path.join(
                "downloads", school_folder, query_task, 'search_results_information_got.xlsx'
        ) not in os.listdir(
            os.path.join("downloads", school_folder, query_task)
        ):
            try:
                crawl.json_to_excel(os.path.join("downloads", school_folder, query_task))
            except:
                continue

    crawl.combine_excel("downloads")
