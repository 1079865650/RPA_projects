# _*_ coding : utf-8 _*_
# @Time : 2023-02-23 13:59
# @Author : wws
# @File : test
# @Project : python基础
import datetime
import os
import re
import redis
import openpyxl
from settings import *
from openpyxl import Workbook


# download the file name to be downloaded to Excel
def write_to_excel(aa, site):
    print("=========filename_list", aa)
    today = datetime.datetime.now().strftime("%Y-%m")
    data_path_splice = os.path.join(origin_path, today)
    if not os.path.exists(data_path_splice):  # create a folder name by month
        os.makedirs(data_path_splice)
    excel_data = os.path.join(data_path_splice, r'data.xlsx')
    if not os.path.exists(excel_data):  # create a excel
        wb = Workbook()
        wb.save(excel_data)
        wb.close()
    wb = openpyxl.load_workbook(excel_data)  # load excel driver
    sheet = wb.active
    avail_row = sheet.max_row  # obtain the maximum row available in Excel
    site_dict = {'EU': [1, 2], 'US': [3, 4], 'JP': [5, 6]}  # write different columns according to different sites
    now_site = site_dict[site]
    for index, data01 in enumerate(aa):
        data = str(data01)
        sheet.cell(row=avail_row + index + 1, column=now_site[0]).value = data  # write raw data
        re_space = re.sub(r'[" "]', '', data)
        sheet.cell(row=avail_row + index + 1, column=now_site[1]).value = re_space  # write data after removing spaces
    wb.save(excel_data)


# read excel,which contains the file name to be downloaded
def read_excel(site):
    site_dict = {'EU': ['A', 'B'], 'US': ['C', 'D'], 'JP': ['E', 'F']}
    site_now = site_dict[site]  # write different columns according to different sites
    excel_path = os.path.join(origin_path, datetime.datetime.now().strftime("%Y-%m") + '\\' + 'data.xlsx')
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    cell_list = []
    for cell in sheet[site_now[1]]:
        if cell.value is not None:
            cell_list.append(cell.value)
    return cell_list


# determined whether to download
def whether_to_download(filename, cell_list, site):
    re_space = re.sub(r'[" "]', '', filename)
    if re_space in cell_list:
        file_path = os.path.join(origin_path, datetime.datetime.now().strftime("%Y-%m") + '\\' + site)
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        return file_path
    return False


# rename the downloaded file according to the site
def rename_file(site, country, us_category=""):
    country_abbreviation = {'Germany': 'DE', 'France': 'FR', 'Italy': 'IT', 'Spain': 'ES', 'United Kingdom': 'UK',
                            'Poland': 'PL', 'Turkey': 'TR', 'Netherlands': 'NL', 'Belgium': 'BE', 'Sweden': 'SE',
                            'Japan': 'JP', 'United States': 'US', 'Mexico': 'MX', 'Canada': 'CA'}
    abbreviation_list = []
    for key in country_abbreviation.keys():
        abbreviation_list.append(country_abbreviation[key])
    prefix = country_abbreviation[country]
    jp_path = os.path.join(origin_path, datetime.datetime.now().strftime("%Y-%m") + '\\' + site)
    g = os.walk(jp_path)
    absolute_path_list = []
    for path, dir_list, file_list in g:
        for file_name in file_list:
            if file_name[0:2] in abbreviation_list:  # determine whether it has been renamed
                continue
            if us_category != "":
                us_category = us_category + " "
            modify_file_name = prefix + " " + us_category + file_name
            modify_path = os.path.join(path, modify_file_name)  # modified file name
            absolute_path = os.path.join(path, file_name)  # raw name
            item = [absolute_path, modify_path]
            absolute_path_list.append(item)
    print(absolute_path_list)
    for i in absolute_path_list:
        a = i[0]
        b = i[1]
        if os.path.exists(b):
            os.remove(b)  # if the file is exists,delete it
        os.rename(a, b)
    return absolute_path_list


# determine the prefix for file naming
def b2b_b2c(value):
    value_split = value.split(" ")
    if 'Unified' in value_split:
        return 'ALL'
    else:
        return 'B2C'


def add(rd_key, rd_value, expired=-1):
    redis_conn = redis.Redis(host=redis_host, port=redis_port, db=redis_db)
    if expired != -1:
        add_result = redis_conn.setex(rd_key, expired, rd_value)
    else:
        add_result = redis_conn.set(rd_key, rd_value)
    redis_conn.close()
    return add_result


def exists(rd_key):
    redis_conn = redis.Redis(host=redis_host, port=redis_port, db=redis_db)
    exist = redis_conn.exists(rd_key)
    redis_conn.close()
    return True if exist > 0 else False


def delete_key(rd_key):
    # redis_conn = connect_redis()
    redis_conn = redis.Redis(host=redis_host, port=redis_port, db=redis_db)
    delete_result = redis_conn.delete(rd_key)
    redis_conn.close()
    return True if delete_result > 0 else False


def add_redis(country):
    add("payment:before:"+country, 1, expiration_time)


def add_redis_download(country):
    add("payment:downloaded:"+country, 1, expiration_time)


def judge_all_b2c(string_value):
    string_split = string_value.split(" ")
    if 'Unified' in string_split:
        return 'ALL'
    return 'B2C'



# add('payment:before:'+'test', 1, 60*60*24*10)
# write_to_excel([1312321, 212321], 'EU')
