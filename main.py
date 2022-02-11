# encoding: utf-8
from crawl import *
from selenium import webdriver

if __name__ == '__main__':    
    ################ Set up parameters here #####################
    default_download_path = "C://Users/bigwh/Downloads" + "/savedrecs.txt"
        # The first string should be the path where your file is downloaded to by default.
        # Most likely, it should be like: "C://Users/usr_nm/Downloads"
    task_list = [ # folder_name, query
            ["results/search_1", "TI=(pFind) AND PY=(2016-2022)"],
            ["results/search_2", "TI=(Attention is All you Need)"]
        ]
        # These are the tasks to be searched
    driver = webdriver.Chrome(
        executable_path='C://Program Files//Google//Chrome//Application//chromedriver.exe'
        # This is the path where you place your chromedriver
    )
    #############################################################
    start_session(driver, task_list, default_download_path)
