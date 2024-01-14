import time
import os
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import wait, expected_conditions


def save_screenshot(driver, prefix, pic_path):
    """Screenshot and save as a png"""
    # 没有被调用，现在
    # paper_id + current_time
    current_time = time.strftime("%Y%m%d-%H%M%S", time.localtime(time.time()))
    driver.save_screenshot(f'{pic_path}{str(prefix)}_{current_time}.png')


# 用于任务记录
class Check:
    """
    用于任务完成的标识和检查的类
    一、该任务所有的工作是否已经完成
    二、该任务系统提供的文件是否已经下载
    三、该搜索结果的子页面是否处理结束
    四、子页面:该搜索结果的子页面是否是否下载
    五、子页面：该搜索结果的子页面信息是否已经提取结束
    """
    # 用于确定任务是否完成
    @staticmethod
    def mark_task_finish_flag(path_school, query_task):
        """Create a flag in the path to mark the task as completed."""
        with open(os.path.join(path_school, query_task, 'completed.flag'), 'w', encoding='utf-8') as f:
            f.write('1')
        return True

    @staticmethod
    def check_task_finish_flag(path_school, query_task):
        """Check if the flag in the path to check if task has been searched."""
        task_path = os.path.join(path_school, query_task)
        return os.path.exists(task_path) and 'completed.flag' in os.listdir(str(task_path))

    # 检查record_sys.txt是否已经下载
    @staticmethod
    def check_task_downloaded_file(path_school, query_task):
        task_path = os.path.join(path_school, query_task)
        return os.path.exists(task_path) and 'record_sys.txt' in os.listdir(str(task_path))

    # 用于确定子页面任务是否完成
    @staticmethod
    def mark_subpage_done(path_school, query_task, current_url):
        subpage_name = get_file_name(current_url)
        with open(
                os.path.join(path_school, query_task, f'{subpage_name}.flag'),
                'a+', encoding='utf-8') as file:
            file.write(subpage_name)
            file.write('\n')

    @staticmethod
    def check_subpage_done(path_school, query_task, current_url):
        subpage_name = get_file_name(current_url)
        task_path = os.path.join(path_school, query_task)
        return os.path.exists(task_path) and f'{subpage_name}.flag' in os.listdir(str(task_path))

    # 用于搜索结果下载记录（每个任务下的每篇搜索文章）
    @staticmethod
    def check_subpage_downloaded(path_school, query_task, current_url):
        subpage_name = get_file_name(current_url)
        file_path = os.path.join(path_school, query_task)
        if (
                os.path.exists(file_path) and f'{subpage_name}.html' in os.listdir(str(file_path))
        ) and (
                os.path.exists(file_path) and f'{subpage_name}.dat' in os.listdir(str(file_path))
        ):
            return True
        else:
            return False

    # 用于搜索结果下载记录（每个任务下的每篇搜索文章）
    @staticmethod
    def mark_get_subpage_selected_info(path_school, query_task, sub_page_url):
        with open(os.path.join(path_school, query_task, 'search_result_record.txt'), 'a+', encoding='utf-8') as file:
            file.write(sub_page_url)
            file.write('\n')

    @staticmethod
    def check_get_subpage_selected_info(path_school, query_task, sub_page_url):
        record_urls = set()
        try:
            with (open(os.path.join(path_school, query_task, 'search_result_record.txt'), 'r', encoding='utf-8') as file):
                for record_url in file.readlines():
                    record_url = record_url.strip("\n")
                    record_urls.add(record_url)
            if sub_page_url in record_urls:
                return True
            else:
                return False
        except FileNotFoundError:
            # 这是第一个处理的网页，所以'search_result_record.txt'还没建立起来
            return False

    @staticmethod
    def check_total_handled(path_school, query_task, n_record):
        record_urls = set()
        with (open(os.path.join(path_school, query_task, 'search_result_record.txt'), 'r', encoding='utf-8') as file):
            for record_url in file.readlines():
                record_url = record_url.strip("\n")
                record_urls.add(record_url)
        if len(record_urls) == n_record:
            return True
        else:
            return False


def roll_down(driver, fold=40):
    """
    Roll down to the bottom of the page to load all results
    """
    # fold 倍数或者次数，就是下移500，重复40次
    for i_roll in range(1, fold + 1):
        time.sleep(0.1)
        driver.execute_script(f'window.scrollTo(0, {i_roll * 500});')


def show_more(driver):
    """
    show more information hided
    """
    # Show all authors et al.
    element = wait.WebDriverWait(driver, 20).until(
        expected_conditions.element_to_be_clickable(
            (
                By.XPATH,
                '//button[@id="HiddenSecTa-showMoreDataButton"]'
             )))
    driver.execute_script("arguments[0].click();", element)
    # 一些页面可以定位，一些页面死活定位不了
    # driver.find_element(
    #     By.XPATH,
    #     '/button[@id="HiddenSecTa-showMoreDataButton"]'
    # ).click()
    try:
        driver.find_element(By.XPATH, '//*[text()="...More"]').click()
    except:
        pass
    try:
        driver.find_element(By.XPATH, '//*[text()=" ...more addresses"]').click()
    except:
        pass


def check_if_human(driver):
    """
    检查是否出现机器人检测
    """
    try:
        driver.find_element(
            By.XPATH,
            '//*[contains(text(),"Please verify you are human to proceed."]'
        )
        raise Exception("机器人检测出来了")
    except:
        try:
            driver.find_element(
                By.XPATH,
                '//*[contains(text(),"我是人类"]'
            )
            raise Exception("机器人检测出来了")
        except:
            pass


def get_file_name(current_url):
    subpage_id = current_url.replace("https://webofscience.clarivate.cn/wos/alldb/full-record/", "")
    pre, suf = subpage_id.split(":")
    return pre + "_" + suf


def decorator(func):
    """
    装饰器，作用就是看是否有返回值，没有就使得返回值取值为
    """
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
        except:
            result = ""
        return result

    return wrapper


@decorator
def get_element_text(driver, match_condition):
    """
    应用装饰器
    """
    info = driver.find_element(
        By.XPATH, match_condition
    ).text
    return info


def close_pendo_windows(driver):
    """Close guiding windows"""
    # Cookies
    try:
        driver.find_element(By.XPATH, '//*[@id="onetrust-accept-btn-handler"]').click()
    except:
        pass
    # "reminder me later"
    try:
        driver.find_element(By.XPATH, '//*[@id="pendo-close-guide-5600f670"]').click()
    except:
        pass
    # "Step 1 of 2"
    try:
        driver.find_element(By.XPATH, '//*[@id="pendo-close-guide-30f847dd"]').click()
    except:
        pass
    # "Got it"
    try:
        driver.find_element(By.XPATH, '//button[contains(@class, "_pendo-button-primaryButton")]').click()
    except:
        pass
    # "No thanks"
    try:
        driver.find_element(By.XPATH, '//button[contains(@class, "_pendo-button-secondaryButton")]').click()
    except:
        pass
    # What was it... I forgot...
    try:
        driver.find_element(By.XPATH, '//span[contains(@class, "_pendo-close-guide")').click()
    except:
        pass
    # Overlay
    try:
        driver.find_element(By.XPATH, '//div[contains(@class, "cdk-overlay-container")]').click()
    except:
        pass
        # Overlay dialog
    try:
        driver.find_element(By.XPATH, '//button[contains(@class, "_pendo-close-guide")]').click()
    except:
        pass
    try:
        driver.find_element(By.XPATH, '//button[contains(text(), "全部允许")]').click()
    except:
        pass


# 显示进度
def show_progress(total, current):
    percent = 100*current/total
    progress = "█" * int(percent/10)
    sys.stdout.write(f"\r{percent:3.0f}%|{progress:<10}| {current}/{total}")
    sys.stdout.flush()


# 遍历文件夹中的所有文件
def list_all_files(path):
    all_file_list = list()
    for root, dirs, files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            print(file_path)
            all_file_list.append(file_path)
    return all_file_list
