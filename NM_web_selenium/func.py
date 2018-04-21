# -*- coding: utf-8 -*-

"""
@version: 
@author: lidan 
@file: func.py
@time: 2018/1/16 16:00
"""
from datetime import datetime,timedelta
import xlrd
import xlwt
import csv
from xlutils.copy import copy
import time
from selenium import webdriver
from selenium.webdriver.support.select import Select
from datetime import datetime, timedelta
import calendar
import requests
import os
import sys
import pandas as pd
reload(sys)
sys.setdefaultencoding('utf8')

class func:

    # 封装xPath等待时间
    def find_element_by_xpath(self, driver, xPath, sleep=0.3):
        time.sleep(sleep)
        obj = driver.find_element_by_xpath(xPath)
        return obj

    # 封装classname等待时间
    def find_element_by_class_name(self, driver, classname, sleep=0.3):
        time.sleep(sleep)
        obj = driver.find_element_by_class_name(classname)
        return obj

    # 封装cssselector等待时间
    def find_element_by_css_selector(self, driver, cssselector, sleep=0.3):
        time.sleep(sleep)
        obj = driver.find_element_by_css_selector(cssselector)
        return obj
    # 读excel
    def read_excel(self,fileAddr,sheet):
        try:
            workbook = xlrd.open_workbook(fileAddr)
            ParamsSheet = workbook.sheet_by_name(sheet)
            return ParamsSheet
        except Exception, e:
            print "读excel异常：" + str(e)

    # 写excel
    def write_excel(self,fileAddr,sheet,row,column,value):
        try:
            workbook = xlrd.open_workbook(fileAddr)
            wb = copy(workbook)
            resultsheet = wb.get_sheet(sheet)
            resultsheet.write(row,column,value)
            wb.save(fileAddr)
        except Exception, e:
            print "写excel异常：" + str(e)

    # 获取日期控件日期所在行列
    # def getDayRowCol(self,date): # date: 2018-01-10
    #     try:
    #         week = datetime.strptime(date, "%Y-%m-%d").weekday()
    #         day = date.split('-')[2]
    #         oneweek = datetime.strptime(date[0:8]+"01", "%Y-%m-%d").weekday()
    #         fullrow = (int(day) + oneweek + 1) / 7
    #         remainder = (int(day) + oneweek + 1) % 7
    #         if remainder == 0:
    #             row = fullrow
    #         else :
    #             row = fullrow + 1
    #         col = week + 2
    #         list = [row,col]
    #     except Exception,e:
    #         print str("日期错误：" + e)
    #     return list

    # 获取日期控件月份所在行列
    def getMonthRowCol(self,date):# date: 2018-01-10
        try:
            month = date.split('-')[1]
            row = int(month) / 4
            remainder1 = int(month) % 4
            if remainder1 == 0:
                row = row
            else :
                row = row + 1
            if row == 1:
                col = int(month)
            elif row == 2:
                col = int(month) - 4
            elif row == 3:
                col = int(month) - 8
            list = [row,col]
        except Exception,e:
            print str("日期错误：" + e)
        return list
    def getDayRowCol(self, date):  # date: 2018-01-10
        try:
            week = datetime.strptime(date, "%Y-%m-%d").weekday()
            day = date.split('-')[2]
            oneweek = datetime.strptime(date[0:8] + "01", "%Y-%m-%d").weekday()
            fullrow = (int(day) + oneweek + 1) / 7
            remainder = (int(day) + oneweek + 1) % 7
            if remainder == 0:
                row = fullrow
            else:
                row = fullrow + 1
            if week == 6:
                col = 1
            else:
                col = week + 2
            list = [row, col]
        except Exception, e:
            print str("日期错误：" + e)
        return list

    # 日期控件选择日期
    def ChooseDate(self, driver, date, CalenderxPath):
        year = date.split('-')[0]
        monthRC = self.getMonthRowCol(date)
        dayRC = self.getDayRowCol(date)
        print dayRC
        self.find_element_by_xpath(driver, CalenderxPath).click()  # 日期控件按钮
        self.find_element_by_xpath(driver, '//*[@id="cc"]/div[1]/div[5]/span').click()  # 点击月份年份 如“二月 2018”
        self.find_element_by_xpath(driver, '//*[@id="cc"]/div[2]/div/div[1]/span[2]/input').clear()
        self.find_element_by_xpath(driver, '//*[@id="cc"]/div[2]/div/div[1]/span[2]/input').send_keys(year)  # 填入年份
        self.find_element_by_xpath(driver, '//*[@id="cc"]/div[2]/div/div[2]/table/tbody/tr[' + str(monthRC[0]) + ']/td[' + str(monthRC[1]) + ']').click()  # 选择月份
        self.find_element_by_xpath(driver, '//*[@id="cc"]/div[2]/table/tbody/tr[' + str(dayRC[0]) + ']/td[' + str(dayRC[1]) + ']').click()  # 日期

    # 下拉框选择时间  selectxPath：下拉框按钮xPath    optionxPath：下拉框内容xPath，截止到最后一个数字前的xPath
    def ChooseTime(self, driver, time, selectxPath, optionxPath):
        if time[0:1] == "0":
            Ctime = time[1:2]
        else:
            Ctime = time[0:2]
        self.find_element_by_xpath(driver, selectxPath).click()
        self.find_element_by_xpath(driver, optionxPath + Ctime + '"]').click()  # 选择开始时间 03:00

    # 下拉框 点击下拉框图标，根据xPath选择相应项
    def Select(self,driver,selectxPath,optionxPath,sleep=0.3):
        time.sleep(sleep)
        self.find_element_by_xpath(driver,selectxPath).click()
        self.find_element_by_xpath(driver,optionxPath).click()

    # 下拉框 可抓到select的xPath  如 tbody/tr/td[1]/select
    def select_by_visible_text(self,driver,xPath,visibletext,sleep=0.3):
        time.sleep(sleep)
        dataNumEveryPageSel = self.find_element_by_xpath(driver,xPath)
        Select(dataNumEveryPageSel).select_by_visible_text(visibletext)

    # 获取table控件某行某列的值
    def getTableCell(self,driver,xPath,row,column):
        xpath = xPath +"/tr[" + str(row) + "]/td[" + str(column) + "]"
        celldata = driver.find_element_by_xpath(xpath).text
        return celldata

    # 打开chrome设置下载文件地址
    def OpenChrome(self,loginPageAddr,download_default_directory = 'C:\Users\Administrator\Downloads',implicitly_waittime = 5):
        # 设置chrome下载文件默认路径
        options = webdriver.ChromeOptions()
        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': download_default_directory}
        options.add_experimental_option('prefs', prefs)
        # 打开浏览器
        driver = webdriver.Chrome(chrome_options=options)
        driver.implicitly_wait(implicitly_waittime)
        driver.maximize_window()
        driver.get(loginPageAddr)
        return driver

    # 登录网管
    def login(self,driver,user,pwd):
        # 登录
        self.find_element_by_xpath(driver, '//*[@id="username"]').send_keys(user)
        self.find_element_by_xpath(driver, '//*[@id="password"]').send_keys(pwd)
        time.sleep(10)
        # self.find_element_by_xpath(driver, '//*[@id="kaptcha"]').send_keys('abcd')
        self.find_element_by_xpath(driver,'/html/body/div[1]/div[2]/div[4]').click()
        if self.isElementExist(driver, 'layui-layer-btn0'):
            self.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定

    # 进入frame，多层frame用","分隔
    def switch_to_frame(self, driver, framestr, sleep = 0.2):
        time.sleep(sleep)
        framearr = framestr.split(',')
        driver.switch_to.default_content()
        for frame in framearr:
            driver.switch_to.frame(frame)

    # 获取当前日期的最近一周的日期
    def getLastWeek(self,d = str(datetime.now()).split()[0]):
        d = datetime.strptime(d, "%Y-%m-%d")
        d2 = str(d - timedelta(days=6)).split()
        return d2[0]

    # 获取当前日期的最近一月的日期
    # def getLastMonth(self,d = str(datetime.now()).split()[0]):
    #     d = datetime.strptime(d, "%Y-%m-%d")
    #     if d.month == 1:
    #         month = 12
    #         year = d.year - 1
    #     else:
    #         month = d.month - 1
    #         year = d.year
    #     monthdays = calendar.monthrange(year, month)[1]
    #     d1 = str(d - timedelta(days = monthdays - 1)).split()
    #     return d1[0]
    def getLastMonth(self, d=str(datetime.now()).split()[0]):
        d = datetime.strptime(d, "%Y-%m-%d")
        if d.month == 1:
            month = 12
            year = d.year - 1
        else:
            month = d.month - 1
            year = d.year
        monthdays = calendar.monthrange(year, d.month)[1]
        d1 = str(d - timedelta(days=monthdays - 1)).split()
        return d1[0]
    # 通过接口获取话务信息统计返回的json串
    def queryChartData(self,driver,url,data):
        cookie = [item["name"] + "=" + item["value"] for item in driver.get_cookies()]
        cookies = '; '.join(cookie)
        # print cookies
        headers = {'Cookie': cookies,
                   # 'Accept':'*/*',
                   # 'Accept-Encoding': 'gzip, deflate',
                   # 'Accept-Language': 'zh-CN,zh;q=0.9',
                   # 'Connection': 'keep-alive',
                   # 'Content-Length': '77',
                   # 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                   # 'Host': '20.21.8.2:8080',
                   # 'Origin': 'ttp://20.21.8.2:8080',
                   # 'Referer': 'http://20.21.8.2:8080/CNMS//background/chartstat/tscCallDur.html?_systemId=20.150.0.40',
                   # 'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36',
                   # 'X-Requested-With': 'XMLHttpRequest'
                   }
        r = requests.post(url, data=data, headers=headers)
        return r.json()

    # 结束进程
    def CloseProcess(self,processStr):
        processArr = processStr.split(',')
        for i in processArr:
            os.system('taskkill /F /IM ' + i + '.exe')

    # 获取两个日期相差多少天
    def getDayDiff(self,date1,date2):
        date1arr = date1.split('-')
        date2arr = date2.split('-')
        d1 = datetime(int(date1arr[0]), int(date1arr[1]), int(date1arr[2]))
        d2 = datetime(int(date2arr[0]), int(date2arr[1]), int(date2arr[2]))
        DayDiff = ((d2 - d1).days) + 1
        return DayDiff

    # 删除包含特定字符命名的文件
    def del_files(self,path,str):
        for root, dirs, files in os.walk(path):
            for name in files:
                if str in name:
                    os.remove(os.path.join(root, name))
                    #print ("Delete File: " + os.path.join(root, name))

    # 查找单元格，获得其行号和列号  cellvalue为想要查找的单元格内容，查找范围是：在worksheets中的第RowRange行和第ColRange列中。
    def GetCellRowAndCol(self,Worksheets,CellValue,RowRange,ColRange):
        kk = "kk"
        for j in range(0, ColRange, 1):
            if Worksheets.cell_value(RowRange,j) == CellValue:
                RowNum = RowRange
                ColNum = j
                return RowNum, ColNum
                break
            elif Worksheets.cell_value(RowRange,j) == "":
                kk = ""
                break
        if kk == "":
            print "未找到内容为："+ CellValue + "的单元格！请停止脚本进行检查！"

    # 判断classname为element元素是否存在
    def isElementExist(self, driver, element):
        flag = True
        try:
            driver.find_element_by_class_name(element)
            return flag
        except:
            flag = False
            return flag

    def second2time(self,t):
        s = int(t) % 60
        h = int(t) / 3600
        m = (int(t) % 3600) / 60
        return str(str(h) + 'h' + str(m) + 'm' + str(s) + 's')

    def new_excel(self,excel_name,sheet_name):
        book = xlwt.Workbook(encoding='utf-8')
        for i in range(len(sheet_name)):
            book.add_sheet(sheet_name[i])
        book.save(excel_name)

        # 秒--> xxhxxmxxs格式
    def secondToHms(self, seconds):
        time = int(seconds)
        h = time / 3600
        m = (time - h * 3600) / 60
        s = time - h * 3600 - m * 60
        hms = str(h) + 'h' + str(m) + 'm' + str(s) + "s"
        return hms

    def csv_to_xlsx_pd(self,csv_name,excel_name):#csv转excel
        csv = pd.read_csv(csv_name, encoding='gbk',index_col = 0 )
        print csv
        csv.to_excel(excel_name, sheet_name='data')

    def log2format(self,log_name,format_log_name):
        # 初始化表头
        header_list = ['Number', 'Time', 'TimeStap', 'ServiceType', 'CallType', 'CrossStation', 'CrossSystem',
                       'ExpectResult', 'CallingAddr', 'CalledAddr', 'ChannelOne', 'ChannelTwo', 'Result', 'Remark',
                       'SpecialType', 'DecideData']
        # 新建excel
        format_book = xlwt.Workbook(encoding='utf-8')
        format_sheet = format_book.add_sheet('format_check_data')
        for col in range(len(header_list)):
            format_sheet.write(0, col, header_list[col])
        # 打开原始结果记录表
        workbook = xlrd.open_workbook(log_name)
        table = workbook.sheet_by_name(u'data')
        row_num = table.nrows
        col_num = table.ncols
        # 获取判断数据列位置
        for col in range(col_num):
            table_head = table.cell(0, col).value
            if table_head == u'判断数据':
                check_num = col
        # 按行整理数据
        for row in range(1, row_num):
            # 分割判断数据列
            dict_data = {}
            check_data = table.cell(row, check_num).value
            dict_data['DecideData'] = check_data
            print check_data
            split_list = check_data.split(u'：')
            print split_list
            dict_data['ServiceType'], lastdata = split_list[0], split_list[1]
            otherdata = lastdata.split('-')
            if dict_data['ServiceType'] == '语音':
                # 按列写入新表格
                dict_data['CallType'], dict_data['CrossStation'], dict_data['CallingAddr'], dict_data['CalledAddr'], \
                dict_data['ExpectResult'] = otherdata[0], otherdata[1], otherdata[2], otherdata[4], otherdata[6]
                dict_data['Result'] = table.cell(row, 4).value
                if len(otherdata) == 8:
                    if '为' in otherdata[7]:
                        dict_data['SpecialType'] = otherdata[7][3:]
                    else:
                        if len(otherdata[7][2:]) <= 2:
                            dict_data['SpecialType'] = otherdata[7][2:] + '呼叫'
                        else:
                            dict_data['SpecialType'] = otherdata[7][2:]
                else:
                    Service = table.cell(row, 2).value
                    if '全呼' in Service:
                        dict_data['SpecialType'] = '全呼'
                    else:
                        dict_data['SpecialType'] = '普通呼叫'
                for key in dict_data:
                    index_num = header_list.index(key)
                    format_sheet.write(row, index_num, dict_data[key])
            else:
                dict_data['Result'] = table.cell(row, 4).value
                # dict_data['DecideData']=otherdata[]
                dict_data['CallType'], dict_data['CallingAddr'], dict_data['CalledAddr'], dict_data['ExpectResult'] =otherdata[0], otherdata[1], otherdata[3], otherdata[5]
                for key in dict_data:
                    index_num = header_list.index(key)
                    format_sheet.write(row, index_num, dict_data[key])
        format_book.save(format_log_name)

    def get_calling_addr(self,Log_data):
        msi_calling = Log_data.drop_duplicates(u'CallingAddr').loc[:,u'CallingAddr'].values
        for i in range(len(msi_calling)):
            msi_calling[i] = int(msi_calling[i])
        print msi_calling
        return msi_calling
    def get_called_addr(self,Log_data):
        called = Log_data.drop_duplicates(u'CalledAddr').loc[:, u'CalledAddr'].values
        for i in range(len(called)):
            called[i] = int(called[i])
        print called
        return called

        # msi_called = Log_data.drop_duplicates(u'CalledAddr').loc[:,u'CalledAddr'].values
        # for i in range(len(msi_called)):
        #     msi_called[i] = int(msi_called[i])
        # for i in range(len(msi_calling)):
        #     msi_called[i] = int(msi_calling[i])
        # individual_msi = msi_called+msi_calling
        # print Log_data.drop_duplicates(u'CallingAddr').loc[:,u'CallingAddr'].values + Log_data.drop_duplicates(u'CalledAddr').loc[:,u'CalledAddr'].values


        # return individual_msi
    def get_group_addr(self,):
        pass
            #
            #
            # individual_msi.append(data_LogData[data_type[1]][u'CallingAddr'].drop_duplicates())
            #
            # individual_msi.append(data_LogData[data_type[1]][u'CalledAddr'].drop_duplicates())

    def charToUnic(self,ch):
        tmp_ch = hex(ord(ch))[2:]
        return "0" * (4 - len(tmp_ch)) + tmp_ch


    def get_indi_count(self,di,type):
        def dict_add(x, y):
            for k, v in y.items():
                if k in x.keys():
                    x[k] += v
                else:
                    x[k] = v
            return x
        msi_num_calling = di.drop_duplicates(u'CallingAddr').loc[:, u'CallingAddr'].values
        for i in range(len(msi_num_calling)):
            msi_num_calling[i] = int(msi_num_calling[i])
        msi_num_called = di.drop_duplicates(u'CalledAddr').loc[:, u'CalledAddr'].values
        for i in range(len(msi_num_called)):
            msi_num_called[i] = int(msi_num_called[i])
        # 取每个主叫号码次数
        data_in_ing_count={}
        data_in_ed_count={}
        for i in range(len(msi_num_calling)):
            data = di[(di[u'CallingAddr'] == unicode(str(msi_num_calling[i])))]
            data_in_ing_count[msi_num_calling[i]] = len(data)
        # 取每个被叫号码次数
        for i in range(len(msi_num_called)):
            data = di[(di[u'CalledAddr'] == unicode(str(msi_num_called[i])))]
            data_in_ed_count[msi_num_called[i]] = len(data)
        if type =='msi':
            end = dict_add(data_in_ing_count, data_in_ed_count)
        else:
            end = data_in_ed_count
        return end