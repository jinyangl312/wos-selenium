# encoding: utf-8
import pathlib
import shutil
import time
import logging
import os
import re
import json
import pandas as pd
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import wait, expected_conditions
import common
from selenium import webdriver

check = common.Check


# 定义开始函数
def start(school, query_task, default_download_path, others_information):

    # 设置浏览器options
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-infobars")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    # 去掉浏览器上的正在受selenium控制
    options.add_argument("--disable-blink-features=AutomationControlled")

    # 看文件路径是否包涵文件名词
    if not default_download_path.endswith("/savedrecs.txt"):
        default_download_path += "/savedrecs.txt"

    # driver.get("https://webofscience.clarivate.cn/")
    # wait_for_login(driver)
    # switch_language_to_eng(driver)

    path_school = os.path.join("downloads", school)

    # 检查路径是否存在，该任务是否已经处理过
    if path_school is not None and check.check_task_finish_flag(path_school, query_task):
        return False

    # 实例化并运行程序
    driver = webdriver.Chrome(options=options)
    # 页面放大
    driver.maximize_window()

    # Search query
    if not search_query(driver, path_school, query_task):
        # Stop if download failed for some reason
        return False

    # 从页面下载部分
    # 判断这个搜索任务是否已经下载过WOS提供的下载文件了
    # 在多线程情况下，下面无法使用，因为无法保证某一线程获得其线程下载的文件
    # 并且由于线程较快，使得文件名变化
    # if not check.check_task_downloaded_file(path_school, query_task):
    #     # Download the search results using inner website function
    #     if not download_search_results(driver, default_download_path):
    #         return False
    #     # Deal with the downloaded search results file(.txt)
    #     if not deal_with_downloaded_file(driver, default_download_path, path_school, query_task):
    #         return False

    # 打开每个页面下载页面和爬取信息部分
    # Deal with records
    if not deal_with_records(driver, path_school, query_task, others_information):
        return False

    # Search completed
    if path_school is not None:
        check.mark_task_finish_flag(path_school, query_task)
    driver.quit()


def wait_for_login(driver):
    """Wait for the user to login if wos cannot be accessed directly."""
    try:
        driver.find_element(By.XPATH, '//div[contains(@class, "shibboleth-login-form")]')
        input('Login before going next...\n')
    except:
        pass


def switch_language_to_eng(driver):
    """Switch language from zh-cn to English."""

    wait.WebDriverWait(driver, 10).until(
        expected_conditions.presence_of_element_located((By.XPATH, '//*[contains(@name, "search-main-box")]')))

    # # 先搞掉一些提示窗口
    common.close_pendo_windows(driver)

    try:
        driver.find_element(By.XPATH, '//*[normalize-space(text())="简体中文"]').click()
        driver.find_element(By.XPATH, '//button[@lang="en"]').click()
    except:
        common.close_pendo_windows(driver)
        driver.find_element(By.XPATH, '//*[normalize-space(text())="简体中文"]').click()
        driver.find_element(By.XPATH, '//button[@lang="en"]').click()


def search_query(driver, path, query_task):
    """
    Go to advanced search page, insert query into search frame and search the query.
    """
    # 看path是否是假的值，如果不是就继续，建立对应文件夹，不然用open with 时就会出错
    if path is not None:
        os.makedirs(path, exist_ok=True)
        logging.info(f"{path}文件夹已经建立")
        second_path = os.path.join(path, f'{query_task}')
        os.makedirs(second_path, exist_ok=True)
        logging.info(f"{path}-{query_task}文件夹已经建立")

    # Close extra windows
    if not len(driver.window_handles) == 1:
        handles = driver.window_handles
        for i_handle in range(len(handles) - 1, 0, -1):  # traverse in reverse order
            # Switch to the window and load the page
            driver.switch_to.window(handles[i_handle])
            driver.close()
        driver.switch_to.window(handles[0])

    # Search query
    driver.get("https://webofscience.clarivate.cn/wos/alldb/advanced-search")
    max_retry = 3
    retry_times = 0
    while True:
        try:
            common.close_pendo_windows(driver)
            # Load the page
            wait.WebDriverWait(driver, 10).until(
                expected_conditions.presence_of_element_located(
                    (By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Clear "]')
                )
            )

            # Clear the field
            driver.find_element(By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Clear "]').click()
            # Insert the query
            driver.find_element(By.XPATH, '//*[@id="advancedSearchInputArea"]').send_keys("{}".format(query_task))
            # Click on the search button
            driver.find_element(By.XPATH,
                                '//span[contains(@class, "mat-button-wrapper") and text()=" Search "]').click()
            break
        except:
            retry_times += 1
            if retry_times > max_retry:
                logging.error("Search exceeded max retries")
                return False
            else:
                # Retry
                logging.debug("Search retrying")

    # Wait for the query page
    try:
        # 根据文章链接判断是否加载成功
        wait.WebDriverWait(driver, 10).until(
            expected_conditions.presence_of_element_located((By.CLASS_NAME, 'title-link')))
    except:
        try:
            # No results 
            driver.find_element(By.XPATH, '//*[text()="No records were found to match your filters"]')
            logging.warning('No records were found to match your filters')
            # Mark as completed
            # 没有搜索结果时的任务标记
            if path is not None:
                check.mark_task_finish_flag(path, query_task)
            return False
        except:
            # Search failed
            driver.find_element(By.XPATH, '//div[contains(@class, "error-code")]')
            logging.error(driver.find_element(By.XPATH, '//div[contains(@class, "error-code")]').text)
            return False
    # Go to the next step
    return True


def download_search_results(driver, default_download_path):
    """
    Export the search results using inner website function. The file is downloaded to default path set for the system.
    """
    max_retry = 3
    retry_times = 0
    while True:
        time.sleep(0)
        common.close_pendo_windows(driver)
        # Not support search for more than 1000 results yet
        assert int(driver.find_element(By.XPATH,
                                       '//span[contains(@class, "end-page")]').text) < 1000, "Sorry, too many results!"
        # File should not exist on default download folder
        # assert not os.path.exists(default_download_path), "File existed on default download folder!"
        try:
            common.close_pendo_windows(driver)
            # Click on "Export"          
            driver.find_element(By.XPATH,
                                '//span[contains(@class, "mat-button-wrapper") and text()=" Export "]').click()
            # Click on "Plain text file"  
            try:
                driver.find_element(By.XPATH,
                                    '//button[contains(@class, "mat-menu-item") and text()=" Plain text file "]'
                                    ).click()
            except:
                driver.find_element(By.XPATH,
                                    '//button[contains(@class, "mat-menu-item") and @aria-label="Plain text file"]'
                                    ).click()
            # Click on "Records from:"
            driver.find_element(By.XPATH, '//*[text()[contains(string(), "Records from:")]]').click()
            # Click on "Export"
            driver.find_element(By.XPATH, '//span[contains(@class, "ng-star-inserted") and text()="Export"]').click()
            # Wait for download to complete
            for retry_download in range(4):
                time.sleep(0.5)
                try:
                    # If there is any "Internal error"
                    wait.WebDriverWait(driver, 2).until(
                        expected_conditions.presence_of_element_located(
                            (By.XPATH, '//div[text()="Server encountered an internal error"]')))
                    driver.find_element(By.XPATH, '//div[text()="Server encountered an internal error"]')
                    driver.find_element(By.XPATH,
                                        '//*[contains(@class, "ng-star-inserted") and text()="Export"]').click()
                except:
                    # 下载成功了就退出循环
                    if os.path.exists(default_download_path):
                        break
            # Download completed
            assert os.path.exists(default_download_path), "File not found!"
            return True
        except:
            retry_times += 1
            if retry_times > max_retry:
                logging.error("Crawl outbound exceeded max retries")
                return False
            else:
                # Retry
                logging.debug("Crawl outbound retrying")
                common.close_pendo_windows(driver)
                # Click on "Cancel"
                try:
                    driver.find_element(By.XPATH,
                                        '//*[contains(@class, "mat-button-wrapper") and text()="Cancel "]').click()
                except:
                    driver.refresh()
                    time.sleep(0)
                wait.WebDriverWait(driver, 10).until(
                    expected_conditions.presence_of_element_located((By.CLASS_NAME, 'title-link')))
            continue


def deal_with_downloaded_file(driver, default_download_path, dst_path, query_task):
    """
    将下载的结果转移至当前路径中
    检查下载文件中的文献条目数目是否等于搜索出来的文献条目数据
    Process the outbound downloaded to the default path set for the system.
    """
    # Move the outbound to dest folder
    assert os.path.exists(default_download_path), "File not found!"
    # 判断目标时路径还是文件
    if pathlib.Path(dst_path).is_dir():
        dst_path = os.path.join(dst_path, f'{query_task}\\record_sys.txt')
    # 该函数功能就是移动
    shutil.move(default_download_path, dst_path)
    logging.debug(f'Outbound saved in {dst_path}')

    # Load the downloaded file (for debug)
    # 检查下载的文献条目数目是否等于网页上显示的搜索文献数目
    with open(dst_path, "r", encoding='utf-8') as f_file:
        n_record_ref = len(re.findall("\nER\n", f_file.read()))
        assert n_record_ref == int("".join(
            driver.find_element(By.XPATH, '//span[contains(@class, "brand-blue")]').text.split(
                ","))), "Records num do not match outbound num"
    return True


def deal_with_records(driver, path, query_task, others_information):
    """
    Open records as new subpages, download or parse subpages according to the setting.
    add a function: 将

    open each search result(record) as a new subpage
    deal with subpages(using function:process_windows();用windows是因为只处理已经打开的subpage,也就相当于桌,所以需要重复调用):
        download all subpages
        get_subpage_inf_wanted(整理页面内所有的信息)
    """
    # init
    # 计算有多少个子网页需要打开
    n_record = int(driver.find_element(By.XPATH, '//span[contains(@class, "brand-blue")]').text)
    n_page = (n_record + 50 - 1) // 50
    assert n_page < 2000, "Too many pages"
    logging.debug(f'{n_record} records found, divided into {n_page} pages')

    # 搜索结果计数
    records_id = 0
    # 创建一个集合
    url_set = set()

    for i_page in range(n_page):
        # 当前浏览器有多少个窗口？等于1，否则出错
        assert len(driver.window_handles) == 1, "Unexpected windows"

        # 先到低，使得所有的summary-record-title-link都展现出来，以便获取元素集
        common.roll_down(driver)

        # Open every record in a new window
        windows_count = 0
        for record in driver.find_elements(By.XPATH, '//a[contains(@data-ta, "summary-record-title-link")]'):
            # 判断该搜索结果（记录）是否已经打开过
            if record.get_attribute("href") in url_set:
                # coz some records have more than 1 href link
                continue
            else:
                url_set.add(record.get_attribute("href"))
            # 新加的一个判断，功能和上面一样的
            # 看这个页面搞完了没
            current_url = record.get_attribute("href")
            if check.check_subpage_done(path, query_task, current_url):
                continue
            # open one pages
            driver.execute_script(f'window.open(\"{record.get_attribute("href")}\");')
            time.sleep(1)
            windows_count += 1
            # 一个条件是是否到10个了。另一个条件是否是5的倍数。
            # 也就是说，只有10,15,20等才会处理
            # 先要多线程的话，好像只能这么写
            if windows_count >= 10 and not windows_count % 5:
                # Save records and close windows
                # 返回值是处理了几个网页
                increment = process_windows(driver, path, query_task, others_information)
                if increment != -1:
                    records_id += increment
                else:
                    return False
                time.sleep(0)
        # 这应该是最后一页，单着的几页。比如53个搜索结果的剩余3个。
        # Save records and close windows
        increment = process_windows(driver, path, query_task, others_information)
        if increment != -1:
            records_id += increment
        else:
            return False
        # Go to the next page
        if i_page + 1 < n_page:
            element = wait.WebDriverWait(driver, 20).until(
                expected_conditions.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//button[contains(@data-ta, "next-page-button")]'
                     )))
            driver.execute_script("arguments[0].click();", element)
    # 根据下载的数据数目是否与查询到的数目对应，确定返回值
    if check.check_total_handled(path, query_task, n_record):
        return True
    else:
        logging.error("Record handled num does equate the num searched!")
        return False


def process_windows(driver, path, query_task, others_information):
    """
    Process all subpages
    records_id: the number of the search result
    path: path of the task
    """
    handles = driver.window_handles
    has_error = False
    for i_handle in range(len(driver.window_handles) - 1, 0, -1):  # traverse in reverse order
        # Switch to the window and load the page
        driver.switch_to.window(handles[i_handle])  # 先打开最后一个
        common.close_pendo_windows(driver)
        current_url = driver.current_url
        time.sleep(1)
        # 展开页面隐藏内容
        common.show_more(driver)
        # 下载页面
        if not check.check_subpage_downloaded(path, query_task, current_url):
            if not download_subpage(driver, path, query_task):
                logging.error(f"Page({current_url}) download mistake!")
                has_error = True
                driver.close()
                continue
        # 获得整理后的信息
        if not check.check_get_subpage_selected_info(path, query_task, current_url):
            if not get_subpage_inf_wanted(driver, path, query_task, others_information):
                logging.error(f"Page({current_url}) information get mistake!")
                has_error = True
                driver.close()
                continue
        driver.close()
        # 标一下，这个页面搞完了
        check.mark_subpage_done(path, query_task, current_url)

    driver.switch_to.window(handles[0])
    # 如果无误的话，返回len(handles)-1，也就是打开了多少个窗口 否则返回-1
    return len(handles) - 1 if not has_error else -1


def download_subpage(driver, path, query_task):
    """
    Download the page to the path
    """
    try:
        # Load the page or throw exception
        wait.WebDriverWait(driver, 10).until(
            expected_conditions.presence_of_element_located((By.XPATH, '//h2[contains(@class, "title")]')))

        current_url = driver.current_url
        subpage_name = common.get_file_name(current_url)

        # Download the record
        with open(os.path.join(path, query_task, f"{subpage_name}.html"), 'w', encoding='utf-8') as file:
            file.write(driver.page_source)
            logging.debug(f'record # {subpage_name} saved in {path}/{query_task}')

        # Download the record in dat
        with open(os.path.join(path, query_task, f"{subpage_name}.dat"), 'w', encoding='utf-8') as file:
            file.write(driver.page_source)
            logging.debug(f'record # {subpage_name} saved in {path}/{query_task}')
        return True
    except:
        return False


def get_subpage_inf_wanted(driver, path,  query_task, others_information):
    """
    爬取信息以下信息：论文名词，期刊，发表时间，作者，单位，类型，
    """
    try:
        # 网页
        current_url = driver.current_url
        # 论文名词
        title_en = driver.find_element(
            By.XPATH, '//*[@id="FullRTa-fullRecordtitle-0" and @lang="en"]'
        ).text.capitalize()
        title_zh_cn = common.get_element_text(
            driver, '//*[@id="FullRTa-fullRecordtitle-0" and @lang="zh-cn"]'
        )
        # 期刊名 无连接型
        # nolink和link不一样，一个是span,一个是a
        try:
            try:
                journal_en = driver.find_element(
                    By.XPATH,
                    '//span[contains(@class, "summary-source-title") and @lang="en"]'
                ).text.capitalize()
                journal_zh_cn = common.get_element_text(
                    driver,
                    '//span[contains(@class, "summary-source-title") and @lang="zh-cn"]'
                )
            except:
                #  期刊名 有连接
                journal_en = driver.find_element(
                    By.XPATH,
                    '//a[contains(@class, "summary-source-title-link") and @lang="en"]'
                ).text.capitalize()
                journal_zh_cn = common.get_element_text(
                    driver,
                    '//a[contains(@class, "summary-source-title-link") and @lang="zh-cn"]'
                )
        except:
            journal_en = "该文没有期刊名称，请注意文章类型（Document Type）"
            journal_zh_cn = ""
        # 作者信息_英文
        authors_info_en = get_authors_info(driver, language="en")
        # 中文相关信息
        authors_info_zh_cn = get_authors_info(driver, language="zh_cn")
        # 出版日期
        published_date = common.get_element_text(driver, '//span[@name="pubdate"]')
        # 检索日期
        indexed_date = common.get_element_text(driver, '//span[@name="indexedDate"]')
        # 文章类型
        document_type = common.get_element_text(driver, '//span[@id="FullRTa-doctype-0"]')
        # volume
        volume = common.get_element_text(driver, '//span[@id="FullRTa-volume"]')
        # Issue
        issue = common.get_element_text(driver,'//span[@id="FullRTa-issue"]')
        # pagenum
        pagenum = common.get_element_text(driver, '//span[@id="FullRTa-pageNo"]')
        # doi
        doi = common.get_element_text(driver, '//span[@id="FullRTa-DOI"]')
        # 摘要——英文
        abstract_en = common.get_element_text(
            driver, '//div[@id="FullRTa-abstract-basic" and @lang = "en"]/p')
        # 摘要——中文
        abstract_zh_cn = common.get_element_text(
            driver, '//div[@id="FullRTa-abstract-basic" and @lang = "zh-cn"]/p')
        # language
        language = common.get_element_text(driver, '//span[@id="HiddenSecTa-language-0"]')
        # cited num
        try:
            cited_num = driver.find_element(
                By.XPATH, '//*[contains(@id, "FullRRPTa-wos-citation-network-times-cited-count-link")]'
            ).text
        except:
            cited_num = 0
        # corresponding author
        corresponding_author_set = set()
        for i in range(5):
            try:
                corresponding_author_elements = wait.WebDriverWait(driver, 10).until(
                    expected_conditions.presence_of_all_elements_located(
                        (By.XPATH, f'//div[@id="FRAiinTa-RepAddrTitle-{i}"]/div/div')
                    )
                )
                for corresponding_author_element in corresponding_author_elements:
                    corresponding_author = corresponding_author_element.find_element(
                        By.XPATH,
                        './/span[@class="value"]'
                    ).text
                    corresponding_author_set.add(corresponding_author)
            except:
                try:
                    corresponding_author = wait.WebDriverWait(driver, 10).until(
                        expected_conditions.presence_of_element_located(
                            (By.XPATH, f'//div[@id="FRAiinTa-RepAddrTitle-{i}"]/div/div/span[@class="value"]')
                        )
                    ).text
                    corresponding_author_set.add(corresponding_author)
                except:
                    break
        corresponding_author_list = list(corresponding_author_set)

        subpage_inf = dict(
            school_id=str(others_information["学校ID"]),
            school=path,
            teacher_id=str(others_information["教师ID"]),
            teacher_name=others_information["姓名"],
            current_url=current_url,
            title=dict(title_en=title_en, title_zh_cn=title_zh_cn),
            journal=dict(journal_en=journal_en, journal_zh_cn=journal_zh_cn),
            authors_info=dict(authors_info_en=authors_info_en, authors_info_zh_cn=authors_info_zh_cn),
            corresponding_author_list=corresponding_author_list,
            cited_num=cited_num,
            published_date=published_date,
            indexed_date=indexed_date,
            volume=volume,
            issue=issue,
            pagenum=pagenum,
            doi=doi,
            document_type=document_type,
            abstract=dict(abstract_en=abstract_en, abstract_zh_cn=abstract_zh_cn),
            language=language,
            time_data=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()),
        )

        # 写入json
        with open(os.path.join(path, query_task,  'search_results_information_got.txt'), 'a', encoding='utf-8') as file:
            json.dump(subpage_inf, file, indent=4, ensure_ascii=False, allow_nan=True)
            file.write("\n")
            logging.debug(f'record #{current_url} dict saved in {path}-{query_task}')
        file.close(),

        # 写入json
        with open(os.path.join(path, query_task, 'search_results_information_got.json'), 'a', encoding='utf-8') as file:
            json.dump(subpage_inf, file, ensure_ascii=False, allow_nan=True)
            file.write("\n")
            logging.debug(f'record #{current_url} dict saved in {path}-{query_task}')
        file.close()

        # mark一下，这个网页爬下来了
        check.mark_get_subpage_selected_info(path, query_task, current_url)
        return True
    except:
        return False


def get_authors_info(driver, language="en"):
    """
    获得作者信息部分太长了，且可重复，所以单独设为一个函数
    """
    authors_info = list()

    author_elements = driver.find_elements(
        By.XPATH,
        '//div[@id="SumAuthTa-MainDiv-author-{lang}"]/span/span[@class="value ng-star-inserted"]'.format(lang=language)
    )
    # find_elements没找到不会报错，只会返回空列表
    # 没有的话就返回空字典，对应的情况是：全英文，可能没有中文
    if not author_elements:
        author_info = dict(
            author_order="",
            author_name_dis="",
            author_name_std="",
            author_addresses=[],
            author_email="",
        )
        authors_info.append(author_info)
        return authors_info
    # 有的话就继续
    author_order = 0
    for author_element in author_elements:
        # 展示名
        author_name_dis = author_element.find_element(
            By.XPATH,
            './/a[@id="SumAuthTa-DisplayName-author-{lang}-{order}"]'.format(order=author_order, lang=language)
        ).text
        # 标准名
        author_name_std = common.get_element_text(
            author_element,
            './/span[@id="SumAuthTa-FrAuthStandard-author-{lang}-{order}"]/span'.format(
                order=author_order, lang=language
            )
        )
        author_addresses = list()

        # 作者对应地址
        try:
            address_elements = author_element.find_elements(
                By.XPATH,
                './/a[contains(@class,"address_link")]'
            )
        except NoSuchElementException:
            address_elements = [author_order]

        for address_element in address_elements:
            # 得到address编号
            # [1]是个list
            try:
                address_order = address_element[0]
            except TypeError:
                address_order = address_element.text.strip()
                address_order = str(address_order).split("[")[1].split("]")[0]
            # 通过编号得到对应的地址
            try:
                address = driver.find_element(
                    By.XPATH,
                    '//*[@id="address_{}"]/span[2]'.format(address_order)
                ).text
                author_addresses.append(address)
            except:
                author_addresses.append("该文未列出地址信息！")
        # 邮箱
        try:
            author_email = driver.find_element(
                By.XPATH,
                '//a[@id="FRAiinTa-AuthRepEmailAddr-{}"]'.format(author_order)
            ).text
        except:
            author_email = "注意：该文章作者邮箱与作者并不对应！"
        author_order += 1
        author_info = dict(
            author_order=author_order,
            author_name_dis=author_name_dis,
            author_name_std=author_name_std,
            author_addresses=author_addresses,
            author_email=author_email,
        )
        authors_info.append(author_info)
    return authors_info


def json_to_excel(file_path):
    """
    将下载的json数据转成excel
    """
    # 从json文件中加载数据
    # 读取JSON数据并解析为Python数据结构
    with open(os.path.join(file_path, 'search_results_information_got.json'), 'r', encoding='utf-8') as file:
        rows = []
        for line in file:
            data = json.loads(line)
            rows.append(data)

    # 将Python数据结构转换为DataFrame
    df = pd.json_normalize(rows)

    # 将DataFrame保存为Excel文件
    df.to_excel(os.path.join(file_path, 'search_results_information_got.xlsx'), index=False)


def combine_excel(path):
    data_list = []
    # os.listdir(".")返回目录中的文件名列表
    for file in common.list_all_files(path):
        # 判断文件名以".xlsx"结尾
        if file.endswith("search_results_information_got.xlsx"):
            # pd.read_excel(filename)读取Excel文件，返回一个DataFrame对象
            # 列表名.append将DataFrame写入列表
            data_list.append(pd.read_excel(file))
        else:
            continue

    # concat合并Pandas数据
    data_all = pd.concat(data_list)
    # 将 DataFrame 保存为 excel 文件
    data_all.to_excel("all_results.xlsx", index=False)


