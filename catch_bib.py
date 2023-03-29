import re
import os
import openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import selenium
import time

import re
import os
import openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import selenium
import time

def read_txt(path):
    papers = []
    with open(path,'r', encoding='utf-8') as f:
        content=f.readlines()
        for c in content:
            paper_name=c.split('.')[1].split('[')[0]
            papers.append(paper_name)
    return papers

def notrobort(wd, File,index):
    """出现人机认证时，激活该部分，先手动通过图片认证"""
    time.sleep(60)  # 用于图片认证的时间 10s
    element = wd.find_element(By.ID, 'gs_hdr_tsi')
    element.send_keys(File[index] + '\n')
    linkElem = wd.find_element(By.LINK_TEXT, '引用')
    linkElem.click()
    linkElem = wd.find_element(By.LINK_TEXT, 'BibTeX')
    linkElem.click()
    time.sleep(2)
    try:
        element = wd.find_element(By.TAG_NAME, 'pre')
        print(element.text)
        time.sleep(2)
        wd.refresh()
        wd.back()
        wd.back()
        wd.back()
        wd.refresh()
        wd.back()

    except selenium.common.exceptions.NoSuchElementException:
        print('将进行第二次人机验证\n')
        time.sleep(60)  # 用于图片认证的时间 10s
        element = wd.find_element(By.TAG_NAME, 'pre')
        print(element.text)
        time.sleep(2)
        wd.refresh()
        wd.back()
        wd.back()
        wd.back()
        wd.refresh()
        wd.back()

def bib(wd, index, l, File):
    with open('bib.txt', 'w+', encoding='utf-8') as f:
        while index < l:
            try:
                wd.get('https://scholar.google.com/')
                wd.maximize_window()
                element = wd.find_element(By.ID, 'gs_hdr_tsi')
                element.send_keys(File[index] + '\n')
                try:
                    linkElem = wd.find_element(By.LINK_TEXT, '引用')
                    linkElem.click()
                    linkElem = wd.find_element(By.LINK_TEXT, 'BibTeX')
                    linkElem.click()
                    element = wd.find_element(By.TAG_NAME, 'pre')
                    print(element.text)
                    f.write('\n' + element.text)
                    time.sleep(2)
                    wd.refresh()
                    wd.back()
                    wd.back()
                    wd.back()
                    index += 1
                except selenium.common.exceptions.NoSuchElementException:
                    index+=1
            except selenium.common.exceptions.NoSuchElementException:
                wd.quit()
                key = input("请按 Y 进行人机验证\n")
                if key == 'Y':
                    print('将进行人机验证\n')
                    option = webdriver.ChromeOptions()
                    option.add_experimental_option("detach", True)
                    wd = webdriver.Chrome(
                        executable_path=r'D:\anaconda\envs\tf1\chromedriver.exe',
                        options=option)
                    wd.implicitly_wait(5)
                    wd.get('https://scholar.google.com/')
                    wd.maximize_window()
                    notrobort(wd, File,index)
                    # f.write('\n' + bib)
        wd.quit()


def bibdownload(path):
    # File = readfilename(path)
    File=read_txt(path)
    l = len(File)
    index = 0
    option = webdriver.ChromeOptions()
    option.add_experimental_option("detach", True)
    wd = webdriver.Chrome(
        executable_path=r'D:\anaconda\envs\tf1\chromedriver.exe',
        options=option)
    wd.implicitly_wait(5)
    wd.get('https://scholar.google.com/')
    wd.maximize_window()
    bib(wd, index, l, File)

    wd = webdriver.Chrome(
        executable_path=r'D:\anaconda\envs\tf1\chromedriver.exe',
        options=option)
    wd.implicitly_wait(5)
    wd.get('https://www.apple.com.cn/')
    wd.maximize_window()
    time.sleep(12)
    wd.quit()


def paperrush(path):
    bibdownload(path)


if __name__ == "__main__":
    path = r'D:\Documents\ref.txt'
    paperrush(path)