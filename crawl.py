# encoding: utf-8
from logging import handlers
import pathlib
import shutil
import time
import logging
import os
import re
import tqdm
from selenium.webdriver.common.by import By
from selenium.webdriver.support import wait, expected_conditions


def wait_for_login(driver):
    '''Wait for the user to login if wos cannot be accessed directly.'''
    try:
        driver.find_element(By.XPATH, '//div[contains(@class, "shibboleth-login-form")]')
        input('Login before going next...\n')
    except:
        pass


def switch_language_to_Eng(driver):
    '''Switch language from zh-cn to English.'''

    wait.WebDriverWait(driver, 10).until(
        expected_conditions.presence_of_element_located((By.XPATH, '//*[contains(@name, "search-main-box")]')))

    close_pendo_windows(driver)
    try:
        driver.find_element(By.XPATH, '//*[normalize-space(text())="简体中文"]').click()
        driver.find_element(By.XPATH, '//button[@lang="en"]').click()
    except:
        close_pendo_windows(driver)
        driver.find_element(By.XPATH, '//*[normalize-space(text())="简体中文"]').click()
        driver.find_element(By.XPATH, '//button[@lang="en"]').click()


def close_pendo_windows(driver):
    '''Close guiding windows'''
    # Cookies
    try:
        driver.find_element(By.XPATH, '//*[@id="onetrust-accept-btn-handler"]').click()
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
        driver.find_element(By.XPATH, '//div[contains(@class, "cdk-overlay-container")').click()
    except:
        pass      


def mark_flag(path):
    '''Create a flag in the path to mark the task as completed.'''
    with open(os.path.join(path, 'completed.flag'), 'w') as f:
        f.write('1')


def check_flag(path):
    '''Check if the flag in the path to check if task has been searched.'''
    return os.path.exists(path) and 'completed.flag' in os.listdir(path) 
    

def search_query(driver, path, query):
    '''Go to advanced search page, insert query into search frame and search the query.'''
    if not path == None:
        os.makedirs(path, exist_ok=True)
        logging.info(path)

    # Close extra windows
    if not len(driver.window_handles) == 1:
        handles = driver.window_handles
        for i_handle in range(len(handles)-1, 0, -1): # traverse in reverse order
            # Switch to the window and load the page
            driver.switch_to.window(handles[i_handle])
            driver.close()
        driver.switch_to.window(handles[0])

    ## Search query
    driver.get("https://www.webofscience.com/wos/alldb/advanced-search")
    max_retry = 3
    retry_times = 0
    while True: 
        try:
            close_pendo_windows(driver)
            # Load the page
            wait.WebDriverWait(driver, 10).until(
                expected_conditions.presence_of_element_located((By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Clear "]')))

            # Clear the field
            driver.find_element(By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Clear "]').click()
            # Insert the query
            driver.find_element(By.XPATH, '//*[@id="advancedSearchInputArea"]').send_keys("{}".format(query))
            # Click on the search button
            driver.find_element(By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Search "]').click()
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
        wait.WebDriverWait(driver, 5).until(
            expected_conditions.presence_of_element_located((By.CLASS_NAME, 'title-link')))
    except:        
        try:
            # No results 
            driver.find_element(By.XPATH, '//*[text()="Your search found no results"]')
            logging.warning(f'Your search found no results')
            # Mark as completed
            if not path == None:
                mark_flag(path)
            return False
        except:
            # Search failed
            driver.find_element(By.XPATH, '//div[contains(@class, "error-code")]')
            logging.error(driver.find_element(By.XPATH, '//div[contains(@class, "error-code")]').text)
            return False
    # Go to the next step
    return True


def download_outbound(driver, default_download_path):
    '''Export the search results as outbound. The file is downloaded to default path set for the system.'''
    max_retry = 3
    retry_times = 0
    while True: 
        close_pendo_windows(driver)
        # Not support search for more than 1000 results yet
        assert int(driver.find_element(By.XPATH, '//span[contains(@class, "end-page")]').text) < 1000, "Sorry, too many results!"
        # File should not exist on default download folder
        assert not os.path.exists(default_download_path), "File existed on default download folder!"
        try:  
            # Click on "Export"          
            driver.find_element(By.XPATH, '//span[contains(@class, "mat-button-wrapper") and text()=" Export "]').click()
            # Click on "Plain text file"  
            try:
                driver.find_element(By.XPATH, '//button[contains(@class, "mat-menu-item") and text()=" Plain text file "]').click()
            except:
                driver.find_element(By.XPATH, '//button[contains(@class, "mat-menu-item") and @aria-label="Plain text file"]').click()
            # Click on "Records from:"
            driver.find_element(By.XPATH, '//*[text()[contains(string(), "Records from:")]]').click()
            # Click on "Export"
            driver.find_element(By.XPATH, '//span[contains(@class, "ng-star-inserted") and text()="Export"]').click()
            # Wait for download to complete
            for retry_download in range(4):
                time.sleep(2)
                try:
                    # If there is any "Internal error"
                    wait.WebDriverWait(driver, 2).until(
                        expected_conditions.presence_of_element_located((By.XPATH, '//div[text()="Server encountered an internal error"]'))) 
                    driver.find_element(By.XPATH, '//div[text()="Server encountered an internal error"]')
                    driver.find_element(By.XPATH, '//*[contains(@class, "ng-star-inserted") and text()="Export"]').click()
                except:
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
                close_pendo_windows(driver)
                # Click on "Cancel"
                try:
                    driver.find_element(By.XPATH, '//*[contains(@class, "mat-button-wrapper") and text()="Cancel "]').click()
                except:
                    driver.refresh()
                    time.sleep(1)
                wait.WebDriverWait(driver, 10).until(
                    expected_conditions.presence_of_element_located((By.CLASS_NAME, 'title-link')))
            continue


def process_outbound(driver, default_download_path, dst_path):
    '''Process the outbound downloaded to the default path set for the system.'''

    # Move the outbound to dest folder
    assert os.path.exists(default_download_path), "File not found!"
    if pathlib.Path(dst_path).is_dir():
        dst_path = os.path.join(dst_path, 'record.txt')
    shutil.move(default_download_path, dst_path)
    logging.debug(f'Outbound saved in {dst_path}')

    # Load the downloaded outbound (for debug)
    with open(dst_path, "r", encoding='utf-8') as f_outbound:
        n_record_ref = len(re.findall("\nER\n", f_outbound.read()))
        assert n_record_ref == int("".join(driver.find_element(By.XPATH, '//span[contains(@class, "brand-blue")]').text.split(","))), "Records num do not match outbound num"
    return True


def download_record(driver, path, records_id):
    '''Download the page to the path'''
    # Load the page or throw exception
    wait.WebDriverWait(driver, 10).until(
        expected_conditions.presence_of_element_located((By.XPATH, '//h2[contains(@class, "title")]')))

    # Download the record
    with open(os.path.join(path, f'record-{records_id}.html'), 'w', encoding='utf-8') as file:
        file.write(driver.page_source)
        logging.debug(f'record #{records_id} saved in {path}')


def process_record(driver, path, records_id):
    '''Parse a page to get certain statistics'''
    # Show all authors and save raw data
    try:
        driver.find_element(By.XPATH, '//button[text()="...More"]').click()                
    except:
        pass
    with open(os.path.join(path, f'record-{records_id}.dat'), 'w', encoding='utf-8') as file:
        file.write(driver.page_source)
        logging.debug(f'record #{records_id} saved in {path}')  


def roll_down(driver, fold = 40):
    '''Roll down to the bottom of the page to load all results'''
    for i_roll in range(1, fold+1):
        time.sleep(0.1)
        driver.execute_script(f'window.scrollTo(0, {i_roll * 500});') 


def save_screenshot(driver, prefix, pic_path):
    """Screenshot and save as a png"""

    # paper_id + current_time
    current_time = time.strftime("%Y%m%d-%H%M%S", time.localtime(time.time()))
    driver.save_screenshot(f'{pic_path}{str(prefix)}_{current_time}.png')


def process_windows(driver, path, records_id):
    '''Process all subpages'''
    handles = driver.window_handles
    has_error = False
    for i_handle in range(len(driver.window_handles)-1, 0, -1): # traverse in reverse order
        # Switch to the window and load the page
        driver.switch_to.window(handles[i_handle])
        close_pendo_windows(driver)
        try:
            download_record(driver, path, records_id)
            process_record(driver, path, records_id)
        except:
            logging.error("Record downloading failed!")
            has_error = True
        records_id += 1
        driver.close()
    driver.switch_to.window(handles[0])
    return len(handles) - 1 if not has_error else -1


def process_records(driver, path):
    '''Open records as new subpages, download or parse subpages according to the setting.'''
    # init
    n_record = int(driver.find_element(By.XPATH, '//span[contains(@class, "brand-blue")]').text)
    n_page = (n_record + 50 - 1) // 50
    assert n_page < 2000, "Too many pages"
    logging.debug(f'{n_record} records found, divided into {n_page} pages')
    
    records_id = 0
    url_set = set()
    for i_page in range(n_page):
        assert len(driver.window_handles) == 1, "Unexpected windows"
        roll_down(driver)
        
        # Open every record in a new window
        windows_count = 0
        for record in driver.find_elements(By.XPATH, '//a[contains(@data-ta, "summary-record-title-link")]'):
            if record.get_attribute("href") in url_set:
                # coz some records have more than 1 href link   
                continue 
            else:
                url_set.add(record.get_attribute("href"))             
            time.sleep(0.5)
            driver.execute_script(f'window.open(\"{record.get_attribute("href")}\");')
            windows_count += 1
            if windows_count >= 10 and not windows_count % 5:
                # Save records and close windows
                increment = process_windows(driver, path, records_id)
                if increment != -1:
                    records_id += increment
                else:
                    return False
                time.sleep(5)
        
        # Save records and close windows
        increment = process_windows(driver, path, records_id)
        if increment != -1:
            records_id += increment
        else:
            return False
        # Go to the next page
        if i_page + 1 < n_page:            
            driver.find_element(By.XPATH, '//mat-icon[contains(@svgicon, "arrowRight")]').click()
    return True


def start_session(driver, task_list, default_download_path):
    '''
    Start the search of all tasks.
    driver: the handle of a selenium.webdriver object
    task_list: the zip of save paths and advanced query strings
    default_download_path: the default path set for the system, for example, C://Downloads/
    '''
    
    # Init
    os.makedirs('logs', exist_ok=True)
    logging.basicConfig(level=logging.INFO,
                    filename=os.getcwd() + '/logs/log' + time.strftime('%Y%m%d%H%M',
                                                time.localtime(time.time())) + '.log',
                    filemode="w", 
                    format="%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s"
                    )

    if not default_download_path.endswith("/savedrecs.txt"):
        default_download_path += "/savedrecs.txt"
    driver.get("https://www.webofscience.com/")
    wait_for_login(driver)
    switch_language_to_Eng(driver)

    # Start Query
    for path, query in tqdm.tqdm(task_list):
        if not path == None and check_flag(path): continue

        # Search query
        if not search_query(driver, path, query):
            # Stop if download failed for some reason
            continue

        # Download the outbound   
        if not download_outbound(driver, default_download_path):
            continue

        # Deal with the outbound   
        if not process_outbound(driver, default_download_path, path):
            continue

        # Deal with records
        if not process_records(driver, path):
            continue

        # Search completed
        if not path == None:
            mark_flag(path)
        
    driver.quit()
