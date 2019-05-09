import datetime
import random
import xlsxwriter
import bs4
from bs4 import BeautifulSoup
from selenium import webdriver
import requests

session_requests = requests.session()
optional_part_indexes = [2, 2, 2, 2, 1]
hpe_col_numbers = [[0, 1, 2, 3], [0, 1, 2, 3], [0, 2, 3],
                   [0, 1, 2, 3], [1, 2, 3], [1, 2, 3]]
selected_col_numbers = []
selected_key_index = -1
pages_list = [
    "https://support.hpe.com/hpsc/doc/public/display?docLocale=en_US&docId=emr_na-a00052032en_us",
    "https://support.hpe.com/hpsc/doc/public/display?docLocale=en_US&docId=emr_na-a00052033en_us",
    "https://support.hpe.com/hpsc/doc/public/display?docLocale=en_US&docId=emr_na-c04315770",
    "https://support.hpe.com/hpsc/doc/public/display?docLocale=en_US&docId=emr_na-a00061523en_us",
    "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c04506831",
    "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c03720096",
    # "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c04315770",
    # "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c04506831",
    # "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c03713800",
    # "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c03720096",
    # "https://support.hpe.com/hpsc/doc/public/display?docId=emr_na-c00305257",
]
hpe_products_list = []
existing_list = []
desktop_agents = [
    'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_1) AppleWebKit/602.2.14 (KHTML, like Gecko) Version/10.0.1 Safari/602.2.14',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0']


def random_headers():
    return {'User-Agent': random.choice(desktop_agents)}


def parse_row(row):
    try:
        cols = row.find_all('td')
        max_cols = selected_col_numbers[len(selected_col_numbers) - 1] + 1
        global hpe_products_list
        if len(cols) >= max_cols:
            row_list = []
            for col_index in selected_col_numbers:
                data = str(cols[col_index].text).replace("\t", "").replace("\n", "").strip()
                if col_index == selected_key_index:
                    row_list.insert(0, data)
                else:
                    row_list.append(data)
            # for row in row_list:
            if not hpe_products_list.__contains__(row_list):
                hpe_products_list.append(row_list)

    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def parse_table(table):
    try:
        t_body = table.find('tbody')
        rows = t_body.find_all('tr')
        for row in rows:
            parse_row(row)
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def parse_page(url):
    try:
        result = session_requests.get(
            url,
            headers=random_headers(),
        )
        print(result.ok, result.status_code)
        if result.ok and result.status_code == 200:
            soup = BeautifulSoup(result.text, "lxml")
            tables = soup.find_all('table', attrs={'class': 'ods_si_table_bordered'})
            print(len(tables))
            for table in tables:
                parse_table(table)

    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        # print(hpe_products_list)
        print(len(hpe_products_list))


def read_existing_numbers():
    try:
        global existing_list
        existing_list = []
        ids_file = open("hpe_models.txt", "r")
        lines = ids_file.readlines()
        for line in lines:
            line = line.replace("\n", "")
            existing_list.append(line.split(","))
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        ids_file.close()


def write_existing_numbers(hpe_list):
    try:
        global existing_list
        ids_file = open("hpe_models.txt", "a")
        for item_row in hpe_list:
            filtered_row = []
            for item in item_row:
                options = str(item).split("/")
                for option_item in options:
                    split_options = option_item.strip(" ").split(" ")
                    for option in split_options:
                        if len(option) < 5 or len(
                                option) > 18 or option == "available" or option == "Unavailable" or option == "NA" or option == "listed" or option == "-":
                            print("***", option)
                        else:
                            filtered_row.append(option)
            if len(filtered_row) > 0:
                if not existing_list.__contains__(filtered_row):
                    existing_list.append(filtered_row)
                    writable_row = ""
                    for option in filtered_row:
                        writable_row += option + ","
                    writable_row = writable_row.strip(",")
                    ids_file.write(writable_row + "\n")
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))
    finally:
        ids_file.close()


def parse_pages_list():
    try:
        for i in range(0, len(pages_list)):
            link = pages_list[i]
            global selected_col_numbers
            selected_col_numbers = hpe_col_numbers[i]
            global selected_key_index
            selected_key_index = optional_part_indexes[i]
            print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), "Going to hit", link)
            parse_page(link)
            print(hpe_products_list)
    except Exception as e:
        print(str(datetime.datetime.now().strftime("[%I:%M:%S %p] -")), str(e))


def parse_part_surfer(url):
    try:
        print("Going To:", url)
        result = session_requests.get(url, headers=random_headers())
        print(result.ok, result.status_code)
        if result.ok and result.status_code == 200:
            soup = BeautifulSoup(result.text, 'lxml')
            main_div = soup.find('div', attrs={'id': 'ctl00_BodyContentPlaceHolder_dvProdinfo'})
            all_tds = main_div.find_all('td')
            for td in all_tds:
                try:
                    part_number = td.find('a').text
                    if part_number not in hpe_products_list and part_number is not None:
                        part_number = part_number.replace("\n", "")
                        if part_number != "":
                            hpe_products_list.append(part_number)
                except:
                    pass

    except Exception as e:
        print(e)


if __name__ == '__main__':
    read_existing_numbers()
    parse_pages_list()
    # parse_part_surfer("http://partsurfer.hpe.com/Search.aspx?SearchText=778456-B21")
    write_existing_numbers(hpe_products_list)
