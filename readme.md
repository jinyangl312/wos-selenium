# wos-selenium

Web of Science spider implemented with selenium.

This project mimic the click of mouse to query WoS, export plain text file and download results automatically. Since the new WoS does not support requests and posts easily with the help from Pendo.io (nice company though), old platforms using scrapy may no longer be used anymore.

You can download the code, set up the config in `main.py` and run `main.py` to test the script. You can also follow the code and descriptions in `demo.ipynb`. You should install selenium and set up chromedriver (or firefox driver, etc.) before running the code.

The logic of this project is: For each task, insert the query and do advanced search; then export the plain text file and move it to the destination path; finally open all results in new windows and download them to the destination path.

This project supports English and Simplified Chinese for the time. For other languages, please change the function in `crawl.py/switch_language_to_Eng` with the help of development tools on your browser. You are welcome to fork this project and do improvements on it. 

jinyangl

2022.2

---


针对新版WoS的爬虫，通过模拟浏览器鼠标点击进行WoS的批量化查询、下载、处理。

使用方法：修改`main.py`中的参数，运行文件。也可以跟随`demo.ipynb`的介绍，以更详细地了解爬虫的代码（~~解决我还没解决的bug~~）。运行前确保selenium和chromedriver（或者firefox driver等）已经安装完毕。

代码的逻辑：对于每个任务，首先模拟输入query进行高级检索，然后下载纯文本文件，移动到path中，再打开所有的结果页面，下载到path中。

至今为止这大概是第一个针对新版WoS的爬虫，希望可以抛砖引玉。代码逻辑应该比较清晰，如果有其他需求可以自行修改，也欢迎与我交流（~~挖坑~~）。

----
更新网址：
https://webofscience.clarivate.cn/wos/alldb/advanced-search
