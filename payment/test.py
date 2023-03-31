# _*_ coding : utf-8 _*_
# @Time : 2023-03-30 10:10
# @Author : wws
# @File : test
# @Project : RPA_projects
import datetime
import os
import re
import openpyxl
from settings import *
from openpyxl import Workbook


def read_excel(site):
    site_dict = {'EU': ['A', 'B'], 'US': ['C', 'D'], 'JP': ['E', 'F']}
    site_now = site_dict[site]
    excel_path = os.path.join(origin_path, datetime.datetime.now().strftime("%Y-%m-%d") + '\\' + 'data.xlsx')
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    cell_list = []
    for cell in sheet[site_now[1]]:
        print(cell.value)
        cell_list.append(cell.value)
    print(cell_list)


def re_re(name_list):
    for i in name_list:
        print(i)
        a = re.sub(r'[" "]', "", str(i))
        print(a)

#
str1 = 'Monthly Unified Summary Report for Feb 1, 2023 00:00 PST - Feb 28, 2023 23:59 PST'
str2 = 'Monthly Unified Transaction Report for Feb 1, 2023 00:00 PST - Feb 28, 2023 23:59 PST'
# name_list = [str1, str2]
# re_re(name_list)
# str3 = str1.upper()
# print(str3)

def rename_file(folder_path):
    country_abbreviation = {'Germany': 'DE', 'France': 'FR', 'Italy': 'IT', 'Spain': 'ES', 'UnitedKingdom': 'UK',
                            'Poland': 'PL', 'Turkey': 'TR', 'Netherlands': 'NL', 'Belgium': 'BE', 'Sweden': 'SE',
                            'Japan': 'JP', 'UnitedStates': 'US', 'Mexico': 'MX', 'Canada': 'CA'}
    g = os.walk(folder_path)
    item_list = []
    for path, dir_list, file_list in g:
        for i in file_list:
            absolute_path = os.path.join(path, i)
            mo_path = str(i).split(" ")
            modify_path = os.path.join(path, mo_path[0] + mo_path[2])
            item = [absolute_path, modify_path]
            item_list.append(item)
    for i in item_list:
        a = i[0]
        b = i[1]
        os.rename(a, b)


# obtain the file name of the file that needs to be moved
def remove_file_to_smb():
    file_path = os.path.join(origin_path, datetime.datetime.now().strftime("%Y-%m"))
    g = os.walk(file_path)
    dir_file = []
    for path, dir_list, file_list in g:
        for i in file_list:
            if i == 'data.xlsx':
                continue
            absolute_path = os.path.join(path, i)
            folder = path.split("\\")[-1]
            item = [folder, absolute_path]
            dir_file.append(item)
    print("=========dir_file folder and filename", dir_file)
    return dir_file


remove_file_to_smb()


















