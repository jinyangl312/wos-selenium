# encoding: utf-8
from crawl import *
from selenium import webdriver

if __name__ == '__main__':    
    ################ Set up parameters here #####################
    default_download_path = "D://Downloads/savedrecs.txt"
    task_list = [
            ["results/search_1", "TS=(pFind)"],
            ["results/search_2", "TI=(Attention is All you Need)"]
        ]
    driver = webdriver.Chrome(
        executable_path='C://Program Files//Google//Chrome//Application//chromedriver.exe'
    )
    #############################################################
    start_session(driver, task_list, default_download_path)
