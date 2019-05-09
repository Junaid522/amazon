import datetime
import os
import time
import xlsxwriter
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup

filter_products_list = []
blacklist_companies = []
whitelist_companies = []

approved_products_list = []
writable_products_list = {}
driver = None
current_search_model = ""
models_list = []
ip_auth_proxies = []
current_proxy_index = 0
max_ip_proxies = 0
proxy_failure_sleep_per_page = 2
per_url_sleep = 2
USE_PROXY = True


def load_ip_auth_proxies():
    try:
        ip_auth_proxies_file = open("ip_auth_proxies.txt", "r")
        lines = ip_auth_proxies_file.readlines()
        for line in lines:
            line = line.replace("\n", "")
            ip_auth_proxies.append(line)
        global max_ip_proxies
        max_ip_proxies = ip_auth_proxies.__len__()
    except Exception as e:
        print(datetime.time.strftime("[%H:%M:%S]"), str(e))
    finally:
        ip_auth_proxies_file.close()


def load_whitelist_companies():
    try:
        global whitelist_companies
        whitelist_companies = []
        white_list_file = open("whitelist.txt", "r")
        lines = white_list_file.readlines()
        for line in lines:
            line = line.replace("\n", "")
            whitelist_companies.append(line)
    except Exception as e:
        print(datetime.time.strftime("[%H:%M:%S]"), str(e))
    finally:
        white_list_file.close()


def load_blacklist_companies():
    try:
        global blacklist_companies
        blacklist_companies = []
        black_list_file = open("blacklist.txt", "r")
        lines = black_list_file.readlines()
        for line in lines:
            line = line.replace("\n", "")
            blacklist_companies.append(line)
    except Exception as e:
        print(datetime.time.strftime("[%H:%M:%S]"), str(e))
    finally:
        black_list_file.close()


def get_new_proxy():
    proxies = None
    if USE_PROXY:
        global current_proxy_index
        if current_proxy_index >= max_ip_proxies:
            current_proxy_index = 0
        if current_proxy_index < max_ip_proxies:
            proxies = ip_auth_proxies[current_proxy_index]
            # proxies = {
            #     "http": ip_auth_proxies[current_proxy_index],
            #     "https": ip_auth_proxies[current_proxy_index],
            # }
        current_proxy_index += 1
        print(proxies)
        return proxies
    return proxies


def search_products(model_num):
    try:
        # driver.get("https://www.amazon.com")
        while not open_url("https://www.amazon.com"):
            print("proxy failed sleeping for", proxy_failure_sleep_per_page, "seconds.")
            time.sleep(proxy_failure_sleep_per_page)
        item_to_search = model_num

        search_field_id = "twotabsearchtextbox"
        search_button_xpath = "//*[@id='nav-search']/form/div[2]/div/input"
        search_field_element = WebDriverWait(driver, 10).until(
            lambda driver: driver.find_element_by_id(search_field_id))
        search_button_element = WebDriverWait(driver, 10).until(
            lambda driver: driver.find_element_by_xpath(search_button_xpath))
        search_field_element.clear()
        search_field_element.send_keys(item_to_search)
        search_button_element.click()
        sleep(3)
        getting_items()
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def getting_items():
    try:
        site_html = driver.page_source
        soup = BeautifulSoup(site_html, 'lxml')
        search_div = soup.find('div', attrs={"class": "s-result-list sg-row"})
        items = search_div.find_all('div', attrs={'class': 'a-section a-spacing-medium'})
        print(len(items))
        for item in items:
            check_product_row(item)
            print("*************************************************************************************")
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        open_filtered_urls()


def check_product_row(item):
    try:
        product_dic = {}
        products_div = item.find('div', class_="a-fixed-left-grid-col a-col-right")
        product_link = products_div.find('a')
        product_owner = products_div.find_all('span', class_="a-size-small a-color-secondary")
        quantity_check_passed = True
        company_check_passed = False
        refurbished_check_pass = True
        # quantity_div = item.find('a', attrs={"class", "a-size-small a-link-normal a-text-normal"})
        # if quantity_div:
        #     quantity_str = str(quantity_div.text).strip(" ")
        #     if quantity_str.__contains__("(") and quantity_str.__contains__(")"):
        #         quantity_tokens = quantity_str.split("(")
        #         if quantity_tokens.__len__() > 1:
        #             quantity_str = quantity_tokens[1]
        #             if quantity_str.__len__() > 0:
        #                 quantity_tokens = quantity_str.split(" ")
        #                 expected_numeric_number = quantity_tokens[0]
        #                 if expected_numeric_number.isnumeric():
        #                     quantity_check_passed = False
        #                     print("found numeric token", expected_numeric_number)
        #                     if int(expected_numeric_number) == 1:
        #                         quantity_check_passed = True
        # print("found 1")
        name = product_link.get('title')
        product_url = product_link.get('href')
        if str(name).upper().__contains__("PACK") or str(name).upper().__contains__("LOT"):
            quantity_check_passed = False
        if str(name).upper().__contains__("(CERTIFIED REFURBISHED)"):
            refurbished_check_pass = False
        print(name)
        print(product_url)
        for owner in product_owner:
            owner_text = str(owner.text).strip(" ").upper()
            if not owner_text == "BY" and not company_check_passed:
                # first check black list companies
                if blacklist_companies.__contains__(owner_text):
                    break
                elif owner_text.__contains__("HEWLETT PACKARD"):
                    company_check_passed = True
                    break
                elif whitelist_companies.__contains__(owner_text):
                    company_check_passed = True
                    break
            if company_check_passed:
                break

        print(refurbished_check_pass, company_check_passed, quantity_check_passed)
        if company_check_passed and quantity_check_passed:
            product_dic = {'name': name, 'product_url': product_url, 'refurbished_check_pass': refurbished_check_pass}
            filter_products_list.append(product_dic)
    except Exception as e:
        product_dic = {}
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        return product_dic


def open_filtered_urls():
    try:
        for filter_product in filter_products_list:
            parse_product_url(filter_product)
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def parse_product_url(filter_product):
    try:
        global approved_products_list
        link = filter_product['product_url']
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), "Going to Open", link)
        # driver.get(link)
        while not open_url(link):
            print("proxy failed sleeping for", proxy_failure_sleep_per_page, "seconds.")
            time.sleep(proxy_failure_sleep_per_page)
        site_html = driver.page_source
        soup = BeautifulSoup(site_html, 'lxml')
        input_asin = soup.find('input', attrs={'id': "ASIN"})
        if input_asin:
            filter_product['ASIN'] = input_asin['value']
            print(filter_product['ASIN'])
            if str(site_html).__contains__(current_search_model):
                approved_products_list.append(filter_product)
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def open_url(url):
    try:
        success = False
        time.sleep(per_url_sleep)
        global driver
        if driver:
            driver.close()
        origin_driver = os.getcwd()
        chrome_driver = str(origin_driver) + "/chromedriver"
        if USE_PROXY:
            chrome_options = webdriver.ChromeOptions()
            current_proxy = get_new_proxy()
            chrome_options.add_argument('--proxy-server=%s' % current_proxy)
            driver = webdriver.Chrome(chrome_driver, options=chrome_options)
        else:
            driver = webdriver.Chrome(chrome_driver)
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] - Request")), url)
        driver.get(url)
        if str(driver.page_source).__contains__("This site can’t be reached"):
            print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), "This site can't be reached")
            success = False
        else:
            success = True
    except Exception as e:
        success = False
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        return success


def write_excel_workbook(file_name, refurbished_check_pass):
    try:
        global writable_products_list
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet()
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})
        row = 0
        worksheet.write(row, 0, "sku", bold)
        worksheet.write(row, 1, "product-id", bold)

        for sku, product_list in writable_products_list.items():
            sku_number = 0
            for product in product_list:
                if product['refurbished_check_pass'] == refurbished_check_pass:
                    row += 1
                    sku_number += 1
                    worksheet.write(row, 0, (sku + " - " + str(sku_number)))
                    worksheet.write(row, 1, product['ASIN'])
                    worksheet.write_number(row, 2, 1)
                    if not refurbished_check_pass:
                        worksheet.write_number(row, 6, 11)
                    worksheet.write(row, 8, "a")
                    worksheet.write_number(row, 9, 2)
                    worksheet.write(row, 10, "Next, Second, Domestic, International")
                    worksheet.write(row, 11, "Y")
                    worksheet.write_number(row, 15, 1)

    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        workbook.close()


def read_model_numbers():
    try:
        global models_list
        models_list = []
        ids_file = open("hpe_models.txt", "r")
        lines = ids_file.readlines()
        for line in lines:
            line = line.replace("\n", "")
            models_list.append(line.split(","))
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        ids_file.close()


def search_model(model_row):
    try:
        global approved_products_list
        approved_products_list = []
        global filter_products_list
        filter_products_list = []
        for model_number in model_row:
            global current_search_model
            current_search_model = model_number  # "646894-001"#AJ738A#655710-B21
            search_products(current_search_model)
        writable_products_list[model_row[0]] = approved_products_list
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def search_all_models():
    try:
        for model_row in models_list:
            search_model(model_row)
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


if __name__ == '__main__':
    try:
        load_whitelist_companies()
        print("WHITE LIST:", whitelist_companies)
        load_blacklist_companies()
        print("BLACK LIST:", blacklist_companies)
        load_ip_auth_proxies()
        read_model_numbers()
        search_all_models()
        if driver:
            driver.close()
    except Exception as es:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(es))
    finally:
        write_excel_workbook('output.xlsx', True)
        write_excel_workbook('refurbished_output.xlsx', False)
