# -*- coding: utf-8 -*-

"""
@version: 
@author: lidan 
@file: EOES.py
@time: 2018/1/3 11:04
"""
from selenium import webdriver
import time
from func import func
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import shutil,os
from datetime import datetime, timedelta
import requests
import pandas as pd
import xlwt
import xlrd
import re
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')
class EOES:
    def __init__(self):
        self.Func = func()
        # 呼叫
        self.CallText = ["Group Call,", "Individual Call,", "Inter-BS,", "Intra-BS,", "Intra-system,", "Inter-system,","Normal Call", "GPS Group Call", "Include Call", "Patch Group Call", "PSTN Call", "Reserved Call","Broadcast Call", "Full-duplex Call", "Emergency Call", "Encrypted Call", "Ambience Listening"]
        # 数据业务
        self.DataServiceText = ["Group Call", "Individual Call", "Individual Short Message Call", "Individual Packet Data Call", "Individual Status Message Call", "Group Short Message Call","Group Packet Data Call", "Group Status Message Call"]
        #  登记总次数
        self.TotalRegistrationCountText = ["Local Registration,", "Local Deregistration,", "Inter-system Registration,", "Inter-system Deregistration,", "Registration,", "Deregistration,"]
        self.CallTypeCombox = ["All", "Outgoing", "Incoming"]  # 呼叫类型下拉框
        self.CallResultCombox = ["All", "Succeeded", "Failed"]  # 呼叫结果下拉框
        self.InterSystemCombox = ["All", "No", "Yes"]  # 是否跨系统下拉框
        self.type = ['mso', 'tsc', 'vpn']  # 信息统计类型
        # 文件名
        self.resfiles = ["System-level Statistics", "BS-level Statistics", "Organization-level Statistics"]
        # 登记结果 "全部 成功 失败" 接口form data信息
        self.RegisterResult = ['4294967295','0','1']

        # 今天，最近一周，最近一月的日期
        self.todaydate = str(datetime.now()).split()[0]
        self.lastmonth = self.Func.getLastMonth()
        self.lastweek = self.Func.getLastWeek()
        self.StartDateArray = [self.todaydate, self.lastweek, self.lastmonth]
        self.objparam = self.Func.read_excel('config.xls', '话务定制_系统基站组织架构信息统计配置')

        self.CallDurationArray = []
        self.CallCountArray = []
        self.DataServiceArray = []
        self.TotalRegCountArray = []

    '''系统，基站，组织架构信息统计 '''
    def System_BS_Organization_Statistics(self):
        self.Func.del_files(os.getcwd(), 'Statistics')
        for column in range(2, 5, 1):
            # 从配置文件获取参数
            configsheet = self.Func.read_excel('config.xls', 'Sheet1')
            loginPageAddr = str(configsheet.cell_value(1, 0))
            username = str(configsheet.cell_value(1, 1))
            password = str(configsheet.cell_value(1, 2))
            driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
            self.Func.login(driver, username, password)  # 登录网管
            time.sleep(1)
            '''新建excel文件用于存储数据'''
            resxlsbook = xlwt.Workbook()
            resxlssheet = resxlsbook.add_sheet('Sheet1')
            resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] +".xls")

            self.CallDuration(driver,resxlsbook, column)  # 总呼叫时长
            self.IndivisualCallDuration(driver,resxlsbook, column)  # 单呼呼叫时长
            self.GroupCallDuration(driver,resxlsbook, column)  # 组呼呼叫时长
            self.CallCount(driver,resxlsbook, column)  # 总呼叫次数
            self.IndivisualCallCount(driver,resxlsbook, column)  # 单呼呼叫次数
            self.GroupCallCount(driver,resxlsbook, column)  # 组呼呼叫次数
            self.DataServiceTraffic(driver, resxlsbook, column)  # 数据业务流量
            self.IndividualDataService(driver, resxlsbook, column)  # 单呼数据业务流量
            self.GroupDataService(driver, resxlsbook, column)  # 组呼数据业务流量
            self.DataServiceCount(driver, resxlsbook, column)  # 数据业务次数
            self.IndividualDataServiceCount(driver, resxlsbook, column)  # 单呼数据业务次数
            self.GroupDataServiceCount(driver, resxlsbook, column)  # 组呼数据业务次数
            self.TotalRegistrationCount(driver, resxlsbook, column)  # 总登记次数
            self.MSResgistrationCount(driver, resxlsbook, column)  # 终端登记次数
            time.sleep(2)
            driver.quit()
    '''呼叫时长'''
    def CallDuration(self, driver,resxlsbook, column):
        print self.resfiles[column - 2] + u"：总呼叫时长"
        CallDurationSheet = resxlsbook.add_sheet(self.objparam.cell_value(30, column - 1), cell_overwrite_ok=True)
        CallDurationSheet.write(0, 0, u"呼叫类别")
        CallDurationSheet.write(0, 1, u"最近一天")
        CallDurationSheet.write(0, 2, u"最近一周")
        CallDurationSheet.write(0, 3, u"最近一月")
        for i in range(0, len(self.CallText), 1):
            CallDurationSheet.write(i + 1, 0, self.CallText[i])
        dateindex = 0
        while dateindex < 3:
            # 通过接口获取呼叫时长数据
            CallDuration1 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '1'})
            self.CallDurationArray.append(CallDuration1[1]) # 组呼
            self.CallDurationArray.append(CallDuration1[0])  # 单呼
            CallDuration2 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '2'})
            self.CallDurationArray.append(CallDuration2[1])  # 跨站
            self.CallDurationArray.append(CallDuration2[0])  # 单站
            CallDuration3 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '3'})
            self.CallDurationArray.append(CallDuration3[0])  # 单系统
            self.CallDurationArray.append(CallDuration3[1])  # 跨系统
            CallDuration4 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '4'})
            self.CallDurationArray.append(CallDuration4[0])  # 普通呼叫
            self.CallDurationArray.append(CallDuration4[8])  # GPS组呼
            self.CallDurationArray.append(CallDuration4[7])  # 包容呼叫
            self.CallDurationArray.append(CallDuration4[9])  # 派接组呼
            self.CallDurationArray.append(CallDuration4[5])  # PSTN呼叫
            self.CallDurationArray.append(CallDuration4[10])  # 预约呼叫
            self.CallDurationArray.append(CallDuration4[4])  # 广播呼叫
            self.CallDurationArray.append(CallDuration4[3])  # 双工呼叫
            self.CallDurationArray.append(CallDuration4[1])  # 紧急呼叫
            self.CallDurationArray.append(CallDuration4[2])  # 加密呼叫
            self.CallDurationArray.append(CallDuration4[6])  # 环境侦听
            for i in range(0, len(self.CallDurationArray), 1):
                CallDurationSheet.write(i + 1, dateindex + 1, self.Func.secondToHms(self.CallDurationArray[i]))
            dateindex = dateindex + 1
            self.CallDurationArray = []
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] +".xls")

    '''单呼呼叫时长'''
    def IndivisualCallDuration(self, driver,resxlsbook, column):
        print self.resfiles[column - 2] + u"：单呼呼叫时长"
        IndivisualCallDurationSheet = resxlsbook.add_sheet(self.objparam.cell_value(31, column - 1), cell_overwrite_ok=True)
        title = [u"一天：all", u"一天：主叫", u"一天：被叫", u"一周：all", u"一周：主叫", u"一周：被叫", u"一月：all", u"一月：主叫", u"一月：被叫"]
        head = ["ISI", "Name", "NE Name/Organization Name", "Organization Node", "Call Duration"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            IndivisualCallDurationSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                IndivisualCallDurationSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能  中文版
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click() # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(4, column - 1), 1).click()  # 点击系统单呼呼叫时长
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for i in range(0, len(self.CallTypeCombox), 1):
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                time.sleep(0.5)
                # 跳转到iframe ID变化的iframe
                iframes = driver.find_elements_by_tag_name('iframe')
                for iframe in iframes:
                    if 'layui-layer-iframe' in iframe.get_property('id'):
                        self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                self.Func.ChooseDate(driver, self.StartDateArray[dateindex],'//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(i) + '"]')  # 呼叫方式
                if self.objparam.cell_value(5, column - 1) == '//*[@id="checkSta"]':
                    self.Func.find_element_by_xpath(driver, self.objparam.cell_value(5, column - 1)).click()  # 基站全选
                elif self.objparam.cell_value(5, column - 1) == '//*[@id="ztree_1_check"]':
                    self.Func.find_element_by_xpath(driver, self.objparam.cell_value(5, column - 1), 1).click()  # 组织架构全选
                self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                if column == 3 and self.CallTypeCombox[i] == "Outgoing":
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                    else:
                        trs = driver.find_elements_by_tag_name('tr')
                        rowcount = len(trs) - 1
                        firstcolumn = 3
                        for NMrow in range(0, rowcount, 1):
                            firstrow = nexttablerow
                            if column == 2:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                            elif column == 3:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                            elif column == 4:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                                IndivisualCallDurationSheet.write(firstrow + 2, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 6))
                            firstcolumn = firstcolumn + 1
                else:
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                    else:
                        trs = driver.find_elements_by_tag_name('tr')
                        rowcount = len(trs) - 1
                        firstcolumn = 3
                        for NMrow in range(0, rowcount, 1):
                            firstrow = nexttablerow
                            if column == 2:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                            elif column == 3:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                            elif column == 4:
                                IndivisualCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                                IndivisualCallDurationSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                                IndivisualCallDurationSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                                IndivisualCallDurationSheet.write(firstrow + 2, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                                IndivisualCallDurationSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 6))
                            firstcolumn = firstcolumn + 1
                nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''组呼呼叫时长'''
    def GroupCallDuration(self,driver,resxlsbook, column):
        print self.resfiles[column - 2] + u"：组呼呼叫时长"
        GroupCallDurationSheet = resxlsbook.add_sheet(self.objparam.cell_value(32, column - 1), cell_overwrite_ok=True)
        title = [u"最近一天", u"最近一周", u"最近一月"]
        head = ["GSI", "Group Name", "NE Name/Organization Name", "Organization Node", "Call Duration"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            GroupCallDurationSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                GroupCallDurationSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(6, column - 1), 1).click()  # 点击系统组呼呼叫时长
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            self.Func.switch_to_frame(driver, 'mainFrame')
            self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
            time.sleep(0.5)
            # 跳转到iframe ID变化的iframe
            iframes = driver.find_elements_by_tag_name('iframe')
            for iframe in iframes:
                if 'layui-layer-iframe' in iframe.get_property('id'):
                    self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
            self.Func.ChooseDate(driver, self.StartDateArray[dateindex], '//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
            self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
            if self.objparam.cell_value(7, column - 1) == '//*[@id="checkSta"]':
                self.Func.find_element_by_xpath(driver, self.objparam.cell_value(5, column - 1)).click()  # 基站全选
            elif self.objparam.cell_value(7, column - 1) == '//*[@id="ztree_1_check"]':
                self.Func.find_element_by_xpath(driver, self.objparam.cell_value(5, column - 1), 1).click()  # 组织架构全选
            self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
            self.Func.switch_to_frame(driver, 'mainFrame')
            if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
            else:
                trs = driver.find_elements_by_tag_name('tr')
                rowcount = len(trs) - 1
                firstcolumn = 3
                for NMrow in range(0, rowcount, 1):
                    firstrow = nexttablerow
                    if column == 2:
                        GroupCallDurationSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                        GroupCallDurationSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                        GroupCallDurationSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                    elif column == 3:
                        GroupCallDurationSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                        GroupCallDurationSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                        GroupCallDurationSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                        GroupCallDurationSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                    elif column == 4:
                        GroupCallDurationSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                        GroupCallDurationSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                        GroupCallDurationSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                        GroupCallDurationSheet.write(firstrow + 2, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                        GroupCallDurationSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 6))
                    firstcolumn = firstcolumn + 1
            nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''呼叫次数'''
    def CallCount(self,driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：呼叫次数"
        CallCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(33, column - 1), cell_overwrite_ok=True)
        CallCountSheet.write(0, 0, u"呼叫类别")
        CallCountSheet.write(0, 1, u"最近一天(次)")
        CallCountSheet.write(0, 2, u"最近一周(次)")
        CallCountSheet.write(0, 3, u"最近一月(次)")
        for i in range(0, len(self.CallText), 1):
            CallCountSheet.write(i + 1, 0, self.CallText[i])
        dateindex = 0
        while dateindex < 3:
            # 通过接口获取呼叫时长数据
            CallCount1 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '5'})
            self.CallCountArray.append(CallCount1[1])  # 组呼
            self.CallCountArray.append(CallCount1[0])  # 单呼
            CallCount2 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2],'item': '6'})
            self.CallCountArray.append(CallCount2[1])  # 跨站
            self.CallCountArray.append(CallCount2[0])  # 单站
            CallCount3 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex], 'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '7'})
            self.CallCountArray.append(CallCount3[0])  # 单系统
            self.CallCountArray.append(CallCount3[1])  # 跨系统
            CallCount4 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex],'EndTime': self.todaydate, 'type': self.type[column - 2],'item': '8'})
            self.CallCountArray.append(CallCount4[0])  # 普通呼叫
            self.CallCountArray.append(CallCount4[8])  # GPS组呼
            self.CallCountArray.append(CallCount4[7])  # 包容呼叫
            self.CallCountArray.append(CallCount4[9])  # 派接组呼
            self.CallCountArray.append(CallCount4[5])  # PSTN呼叫
            self.CallCountArray.append(CallCount4[10])  # 预约呼叫
            self.CallCountArray.append(CallCount4[4])  # 广播呼叫
            self.CallCountArray.append(CallCount4[3])  # 双工呼叫
            self.CallCountArray.append(CallCount4[1])  # 紧急呼叫
            self.CallCountArray.append(CallCount4[2])  # 加密呼叫
            self.CallCountArray.append(CallCount4[6])  # 环境侦听
            for i in range(0, len(self.CallCountArray), 1):
                CallCountSheet.write(i + 1, dateindex + 1, self.CallCountArray[i])
            dateindex = dateindex + 1
            self.CallCountArray = []
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''单呼呼叫次数'''
    def IndivisualCallCount(self,driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：单呼呼叫次数"
        IndivisualCallCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(34, column - 1), cell_overwrite_ok=True)
        title = [u"一天：all：all", u"一天：all：成功", u"一天：all：失败", u"一天：主叫：all", u"一天：主叫：成功", u"一天：主叫：失败", u"一天：被叫：all",
                      u"一天：被叫：成功", u"一天：被叫：失败", u"一周：all：all", u"一周：all：成功", u"一周：all：失败", u"一周：主叫：all", u"一周：主叫：成功",
                      u"一周：主叫：失败", u"一周：被叫：all", u"一周：被叫：成功", u"一周：被叫：失败", u"一月：all：all", u"一月：all：成功", u"一月：all：失败",
                      u"一月：主叫：all", u"一月：主叫：成功", u"一月：主叫：失败", u"一月：被叫：all", u"一月：被叫：成功", u"一月：被叫：失败"]
        head = ["ISI", "Name", "NE Name/Organization Name", "Organization Node", "Call Count"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            IndivisualCallCountSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                IndivisualCallCountSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(10, column - 1), 1).click()  # 点击系统单呼呼叫次数
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for i in range(0, len(self.CallTypeCombox), 1):
                for j in range(0, len(self.CallResultCombox), 1):
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                    time.sleep(0.5)
                    # 跳转到iframe ID变化的iframe
                    iframes = driver.find_elements_by_tag_name('iframe')
                    for iframe in iframes:
                        if 'layui-layer-iframe' in iframe.get_property('id'):
                            self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                    self.Func.ChooseDate(driver, self.StartDateArray[dateindex],'//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                    self.Func.ChooseDate(driver, self.todaydate,'//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(i) + '"]')  # 呼叫方式
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[10]/td/div/span/span/span','//*[@id="_easyui_combobox_i2_' + str(j) + '"]')  # 呼叫结果
                    if self.objparam.cell_value(11, column - 1) == '//*[@id="checkSta"]':
                        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(11, column - 1)).click()  # 基站全选
                    elif self.objparam.cell_value(11, column - 1) == '//*[@id="ztree_1_check"]':
                        self.Func.find_element_by_xpath(driver,self.objparam.cell_value(11, column - 1)).click()  # 组织架构全选
                    self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                    else:
                        trs = driver.find_elements_by_tag_name('tr')
                        rowcount = len(trs) - 1
                        firstcolumn = 3
                        for NMrow in range(0, rowcount, 1):
                            firstrow = nexttablerow
                            if column == 2:
                                IndivisualCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndivisualCallCountSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndivisualCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                            elif column == 3:
                                IndivisualCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndivisualCallCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndivisualCallCountSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                                IndivisualCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                            elif column == 4:
                                IndivisualCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndivisualCallCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndivisualCallCountSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                                IndivisualCallCountSheet.write(firstrow + 2, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                                IndivisualCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 6))
                            firstcolumn = firstcolumn + 1
                    nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''组呼呼叫次数'''
    def GroupCallCount(self,driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：组呼呼叫次数"
        GroupCallCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(35, column - 1), cell_overwrite_ok=True)
        title = [u"一天：all", u"一天：成功", u"一天：失败", u"一周：all", u"一周：成功", u"一周：失败", u"一月：all", u"一月：成功", u"一月：失败"]
        head = ["GSI", "Group Name", "NE Name/Organization Name", "Organization Node", "Call Count"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            GroupCallCountSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                GroupCallCountSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(12, column - 1), 1).click()  # 点击系统组呼呼叫次数
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for j in range(0, len(self.CallResultCombox), 1):
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                time.sleep(0.5)
                # 跳转到iframe ID变化的iframe
                iframes = driver.find_elements_by_tag_name('iframe')
                for iframe in iframes:
                    if 'layui-layer-iframe' in iframe.get_property('id'):
                        self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                self.Func.ChooseDate(driver, self.StartDateArray[dateindex],'//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span','//*[@id="_easyui_combobox_i1_' + str(j) + '"]')  # 呼叫结果
                if self.objparam.cell_value(11, column - 1) == '//*[@id="checkSta"]':
                    self.Func.find_element_by_xpath(driver,self.objparam.cell_value(13, column - 1)).click()  # 基站全选
                elif self.objparam.cell_value(11, column - 1) == '//*[@id="ztree_1_check"]':
                    self.Func.find_element_by_xpath(driver,self.objparam.cell_value(13, column - 1)).click()  # 组织架构全选
                self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                self.Func.switch_to_frame(driver, 'mainFrame')
                if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                    self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                else:
                    trs = driver.find_elements_by_tag_name('tr')
                    rowcount = len(trs) - 1
                    firstcolumn = 3
                    for NMrow in range(0, rowcount, 1):
                        firstrow = nexttablerow
                        if column == 2:
                            GroupCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                            GroupCallCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                        elif column == 3:
                            GroupCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                            GroupCallCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupCallCountSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                            GroupCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                        elif column == 4:
                            GroupCallCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                            GroupCallCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupCallCountSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                            GroupCallCountSheet.write(firstrow + 2, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                            GroupCallCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 6))
                        firstcolumn = firstcolumn + 1
                nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''数据业务流量'''
    def DataServiceTraffic(self,driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：数据业务流量"
        DataServiceTrafficSheet = resxlsbook.add_sheet(self.objparam.cell_value(36, column - 1), cell_overwrite_ok=True)
        DataServiceTrafficSheet.write(0, 0, u"呼叫类别")
        DataServiceTrafficSheet.write(0, 1, u"最近一天(KB)")
        DataServiceTrafficSheet.write(0, 2, u"最近一周(KB)")
        DataServiceTrafficSheet.write(0, 3, u"最近一月(KB)")
        for i in range(0, len(self.DataServiceText), 1):
            DataServiceTrafficSheet.write(i + 1, 0, self.DataServiceText[i])
        dateindex = 0
        while dateindex < 3:
            # 通过接口获取呼叫时长数据
            DataServiceTraffic1 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1), 'BeginTime': self.StartDateArray[dateindex],'EndTime': self.todaydate, 'type': self.type[column - 2],'item': '13'})
            self.DataServiceArray.append(DataServiceTraffic1[1])  # 组呼
            self.DataServiceArray.append(DataServiceTraffic1[0])  # 单呼
            DataServiceTraffic2 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),{'_systemId': self.objparam.cell_value(45, 1),'BeginTime': self.StartDateArray[dateindex],'EndTime': self.todaydate, 'type': self.type[column - 2], 'item': '16'})
            self.DataServiceArray.append(DataServiceTraffic2[0])  # 短消息单呼
            self.DataServiceArray.append(DataServiceTraffic2[1])  # 分组消息单呼
            self.DataServiceArray.append(DataServiceTraffic2[2])  # 状态消息单呼
            self.DataServiceArray.append(DataServiceTraffic2[3])  # 短消息组呼
            self.DataServiceArray.append(DataServiceTraffic2[4])  # 分组消息组呼
            self.DataServiceArray.append(DataServiceTraffic2[5])  # 状态消息组呼

            for i in range(0, len(self.DataServiceArray), 1):
                DataServiceTrafficSheet.write(i + 1, dateindex + 1, self.DataServiceArray[i])
            dateindex = dateindex + 1
            self.DataServiceArray = []
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''单呼数据业务流量'''
    def IndividualDataService(self,driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：单呼数据业务流量"
        IndividualDataServiceSheet = resxlsbook.add_sheet(self.objparam.cell_value(37, column - 1), cell_overwrite_ok=True)
        title = [u"一天：all", u"一天：主叫", u"一天：被叫", u"一周：all", u"一周：主叫", u"一周：被叫", u"一月：all", u"一月：主叫", u"一月：被叫"]
        head = ["ISI", "Name", "NE Name/Organization Name", "Organization Node", "Traffic (KB)"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            IndividualDataServiceSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                IndividualDataServiceSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(16, column - 1), 1).click()  # 点击系统单呼数据业务流量
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for i in range(0, len(self.CallTypeCombox), 1):
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                time.sleep(0.5)
                # 跳转到iframe ID变化的iframe
                iframes = driver.find_elements_by_tag_name('iframe')
                for iframe in iframes:
                    if 'layui-layer-iframe' in iframe.get_property('id'):
                        self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                self.Func.ChooseDate(driver, self.StartDateArray[dateindex], '//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(i) + '"]')  # 呼叫方式
                if self.objparam.cell_value(17, column - 1) == '//*[@id="checkSta"]':
                    self.Func.find_element_by_xpath(driver, self.objparam.cell_value(17, column - 1)).click()  # 基站全选
                elif self.objparam.cell_value(17, column - 1) == '//*[@id="ztree_1_check"]':
                    self.Func.find_element_by_xpath(driver, self.objparam.cell_value(17, column - 1), 1).click()  # 组织架构全选
                self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                self.Func.switch_to_frame(driver, 'mainFrame')
                if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                    self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                else:
                    trs = driver.find_elements_by_tag_name('tr')
                    rowcount = len(trs) - 1
                    firstcolumn = 3
                    for NMrow in range(0, rowcount, 1):
                        firstrow = nexttablerow
                        if column == 2:
                            IndividualDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                            IndividualDataServiceSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                            IndividualDataServiceSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                        elif column == 3:
                            IndividualDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                            IndividualDataServiceSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 3))
                            IndividualDataServiceSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                            IndividualDataServiceSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                        elif column == 4:
                            IndividualDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                            IndividualDataServiceSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                            IndividualDataServiceSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                            IndividualDataServiceSheet.write(firstrow + 2, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                            IndividualDataServiceSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,6))
                        firstcolumn = firstcolumn + 1
                nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''组呼数据业务流量'''
    def GroupDataService(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：组呼数据业务流量"
        GroupDataServiceSheet = resxlsbook.add_sheet(self.objparam.cell_value(38, column - 1),cell_overwrite_ok=True)
        title = [u"最近一天", u"最近一周", u"最近一月"]
        head = ["GSI", "Group Name", "NE Name/Organization Name", "Organization Node", "Traffic (KB)"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            GroupDataServiceSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                GroupDataServiceSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(18, column - 1), 1).click()  # 点击系统组呼数据业务流量
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            self.Func.switch_to_frame(driver, 'mainFrame')
            self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
            time.sleep(0.5)
            # 跳转到iframe ID变化的iframe
            iframes = driver.find_elements_by_tag_name('iframe')
            for iframe in iframes:
                if 'layui-layer-iframe' in iframe.get_property('id'):
                    self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
            self.Func.ChooseDate(driver, self.StartDateArray[dateindex], '//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
            self.Func.ChooseDate(driver, self.todaydate,'//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
            if self.objparam.cell_value(19, column - 1) == '//*[@id="checkSta"]':
                self.Func.find_element_by_xpath(driver, self.objparam.cell_value(19, column - 1)).click()  # 基站全选
            elif self.objparam.cell_value(19, column - 1) == '//*[@id="ztree_1_check"]':
                self.Func.find_element_by_xpath(driver, self.objparam.cell_value(19, column - 1), 1).click()  # 组织架构全选
            self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
            self.Func.switch_to_frame(driver, 'mainFrame')
            if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
            else:
                trs = driver.find_elements_by_tag_name('tr')
                rowcount = len(trs) - 1
                firstcolumn = 3
                for NMrow in range(0, rowcount, 1):
                    firstrow = nexttablerow
                    if column == 2:
                        GroupDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                        GroupDataServiceSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                        GroupDataServiceSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                    elif column == 3:
                        GroupDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                        GroupDataServiceSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                        GroupDataServiceSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                        GroupDataServiceSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                    elif column == 4:
                        GroupDataServiceSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                        GroupDataServiceSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                        GroupDataServiceSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                        GroupDataServiceSheet.write(firstrow + 2, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                        GroupDataServiceSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,6))
                    firstcolumn = firstcolumn + 1
            nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''数据业务次数'''
    def DataServiceCount(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：数据业务次数"
        DataServiceCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(39, column - 1), cell_overwrite_ok=True)
        DataServiceCountSheet.write(0, 0, u"呼叫类别")
        DataServiceCountSheet.write(0, 1, u"最近一天(次)")
        DataServiceCountSheet.write(0, 2, u"最近一周(次)")
        DataServiceCountSheet.write(0, 3, u"最近一月(次)")
        for i in range(0, len(self.DataServiceText), 1):
            DataServiceCountSheet.write(i + 1, 0, self.DataServiceText[i])
        dateindex = 0
        while dateindex < 3:
            # 通过接口获取呼叫时长数据
            DataServiceCount1 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),
                                                           {'_systemId': self.objparam.cell_value(45, 1),
                                                            'BeginTime': self.StartDateArray[dateindex],
                                                            'EndTime': self.todaydate, 'type': self.type[column - 2],
                                                            'item': '17'})
            self.DataServiceArray.append(DataServiceCount1[1])  # 组呼
            self.DataServiceArray.append(DataServiceCount1[0])  # 单呼
            DataServiceCount2 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),
                                                           {'_systemId': self.objparam.cell_value(45, 1),
                                                            'BeginTime': self.StartDateArray[dateindex],
                                                            'EndTime': self.todaydate, 'type': self.type[column - 2],
                                                            'item': '20'})
            self.DataServiceArray.append(DataServiceCount2[0])  # 短消息单呼
            self.DataServiceArray.append(DataServiceCount2[1])  # 分组消息单呼
            self.DataServiceArray.append(DataServiceCount2[2])  # 状态消息单呼
            self.DataServiceArray.append(DataServiceCount2[3])  # 短消息组呼
            self.DataServiceArray.append(DataServiceCount2[4])  # 分组消息组呼
            self.DataServiceArray.append(DataServiceCount2[5])  # 状态消息组呼

            for i in range(0, len(self.DataServiceArray), 1):
                DataServiceCountSheet.write(i + 1, dateindex + 1, self.DataServiceArray[i])
            dateindex = dateindex + 1
            self.DataServiceArray = []
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''单呼数据业务次数'''
    def IndividualDataServiceCount(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：单呼数据业务次数"
        IndividualDataServiceCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(40, column - 1),cell_overwrite_ok=True)
        title = [u"一天：all：all", u"一天：all：成功", u"一天：all：失败", u"一天：主叫：all", u"一天：主叫：成功", u"一天：主叫：失败", u"一天：被叫：all",
                      u"一天：被叫：成功", u"一天：被叫：失败", u"一周：all：all", u"一周：all：成功", u"一周：all：失败", u"一周：主叫：all", u"一周：主叫：成功",
                      u"一周：主叫：失败", u"一周：被叫：all", u"一周：被叫：成功", u"一周：被叫：失败", u"一月：all：all", u"一月：all：成功", u"一月：all：失败",
                      u"一月：主叫：all", u"一月：主叫：成功", u"一月：主叫：失败", u"一月：被叫：all", u"一月：被叫：成功", u"一月：被叫：失败"]
        head = ["ISI", "Name", "NE Name/Organization Name", "Organization Node", "Call Count"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            IndividualDataServiceCountSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                IndividualDataServiceCountSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(22, column - 1), 1).click()  # 点击系统单呼数据业务次数
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for i in range(0, len(self.CallTypeCombox), 1):
                for j in range(0, len(self.CallResultCombox), 1):
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                    time.sleep(0.5)
                    # 跳转到iframe ID变化的iframe
                    iframes = driver.find_elements_by_tag_name('iframe')
                    for iframe in iframes:
                        if 'layui-layer-iframe' in iframe.get_property('id'):
                            self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                    self.Func.ChooseDate(driver, self.StartDateArray[dateindex], '//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                    self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(i) + '"]')  # 呼叫方式
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[10]/td/div/span/span/span','//*[@id="_easyui_combobox_i2_' + str(j) + '"]')  # 呼叫结果
                    if self.objparam.cell_value(23, column - 1) == '//*[@id="checkSta"]':
                        self.Func.find_element_by_xpath(driver,self.objparam.cell_value(23, column - 1)).click()  # 基站全选
                    elif self.objparam.cell_value(23, column - 1) == '//*[@id="ztree_1_check"]':
                        self.Func.find_element_by_xpath(driver,self.objparam.cell_value(23, column - 1)).click()  # 组织架构全选
                    self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                    else:
                        trs = driver.find_elements_by_tag_name('tr')
                        rowcount = len(trs) - 1
                        firstcolumn = 3
                        for NMrow in range(0, rowcount, 1):
                            firstrow = nexttablerow
                            if column == 2:
                                IndividualDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndividualDataServiceCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndividualDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                            elif column == 3:
                                IndividualDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndividualDataServiceCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndividualDataServiceCountSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                                IndividualDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                            elif column == 4:
                                IndividualDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                                IndividualDataServiceCountSheet.write(firstrow, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                                IndividualDataServiceCountSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                                IndividualDataServiceCountSheet.write(firstrow + 2, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                                IndividualDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 6))
                            firstcolumn = firstcolumn + 1
                    nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''组呼数据业务次数'''
    def GroupDataServiceCount(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：组呼数据业务次数"
        GroupDataServiceCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(41, column - 1),cell_overwrite_ok=True)
        title = [u"一天：all", u"一天：成功", u"一天：失败", u"一周：all", u"一周：成功", u"一周：失败", u"一月：all", u"一月：成功", u"一月：失败"]
        head = ["GSI", "Group Name", "NE Name/Organization Name", "Organization Node", "Call Count"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            GroupDataServiceCountSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                GroupDataServiceCountSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(24, column - 1), 1).click()  # 点击系统组呼数据业务次数
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for j in range(0, len(self.CallResultCombox), 1):
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                time.sleep(0.5)
                # 跳转到iframe ID变化的iframe
                iframes = driver.find_elements_by_tag_name('iframe')
                for iframe in iframes:
                    if 'layui-layer-iframe' in iframe.get_property('id'):
                        self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                self.Func.ChooseDate(driver, self.StartDateArray[dateindex],'//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                self.Func.ChooseDate(driver, self.todaydate,'//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(j) + '"]')  # 呼叫结果
                if self.objparam.cell_value(25, column - 1) == '//*[@id="checkSta"]':
                    self.Func.find_element_by_xpath(driver,self.objparam.cell_value(25, column - 1)).click()  # 基站全选
                elif self.objparam.cell_value(25, column - 1) == '//*[@id="ztree_1_check"]':
                    self.Func.find_element_by_xpath(driver,self.objparam.cell_value(25, column - 1)).click()  # 组织架构全选
                self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                self.Func.switch_to_frame(driver, 'mainFrame')
                if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                    self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                else:
                    trs = driver.find_elements_by_tag_name('tr')
                    rowcount = len(trs) - 1
                    firstcolumn = 3
                    for NMrow in range(0, rowcount, 1):
                        firstrow = nexttablerow
                        if column == 2:
                            GroupDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 2))
                            GroupDataServiceCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 4))
                        elif column == 3:
                            GroupDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                            GroupDataServiceCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupDataServiceCountSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                            GroupDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 5))
                        elif column == 4:
                            GroupDataServiceCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 2))
                            GroupDataServiceCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 3))
                            GroupDataServiceCountSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1, 4))
                            GroupDataServiceCountSheet.write(firstrow + 2, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                            GroupDataServiceCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 6))
                        firstcolumn = firstcolumn + 1
                nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''登记总次数'''
    def TotalRegistrationCount(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：登记总次数"
        TotalRegistrationCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(42, column - 1),cell_overwrite_ok=True)
        TotalRegistrationCountSheet.write(0, 0, u"呼叫类别")
        header = [u"一天：all", u"一天：成功", u"一天：失败", u"一周：all", u"一周：成功", u"一周：失败", u"一月：all", u"一月：成功", u"一月：失败"]
        for i in range(0, len(header), 1):
            TotalRegistrationCountSheet.write(0, i + 1, header[i])
        for i in range(0, len(self.TotalRegistrationCountText),1):
            TotalRegistrationCountSheet.write(i + 1, 0, self.TotalRegistrationCountText[i])
        fortime = 2
        for dateindex in range(0, 3, 1):
            for n in range(0, len(self.RegisterResult), 1):
                # 通过接口获取呼叫时长数据
                TotalRegistrationCount1 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),
                                                             {'_systemId': self.objparam.cell_value(45, 1),
                                                              'BeginTime': self.StartDateArray[dateindex],
                                                              'EndTime': self.todaydate, 'type': self.type[column - 2],
                                                              'RegisterResult': self.RegisterResult[n], 'item': '25'})
                self.TotalRegCountArray.append(TotalRegistrationCount1[0])  # 本地登记
                self.TotalRegCountArray.append(TotalRegistrationCount1[1])  # 本地去登记
                self.TotalRegCountArray.append(TotalRegistrationCount1[2])  # 跨系统登记
                self.TotalRegCountArray.append(TotalRegistrationCount1[3])  # 跨系统去登记
                TotalRegistrationCount2 = self.Func.queryChartData(driver, self.objparam.cell_value(44, 1),
                                                             {'_systemId': self.objparam.cell_value(45, 1),
                                                              'BeginTime': self.StartDateArray[dateindex],
                                                              'EndTime': self.todaydate, 'type': self.type[column - 2],
                                                              'RegisterResult': self.RegisterResult[n], 'item': '27'})
                self.TotalRegCountArray.append(TotalRegistrationCount2[0])  # 登记
                self.TotalRegCountArray.append(TotalRegistrationCount2[1])  # 去登记


                for i in range(0, len(self.TotalRegCountArray), 1):
                    TotalRegistrationCountSheet.write(i + 1, fortime - 1, self.TotalRegCountArray[i])
                fortime = fortime + 1
                self.TotalRegCountArray = []
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''终端登记次数'''
    def MSResgistrationCount(self, driver, resxlsbook, column):
        print self.resfiles[column - 2] + u"：终端登记次数"
        MSResgistrationCountSheet = resxlsbook.add_sheet(self.objparam.cell_value(43, column - 1),cell_overwrite_ok=True)
        title = [u"一天：all：all", u"一天：all：成功", u"一天：all：失败", u"一天：不跨：all", u"一天：不跨：成功", u"一天：不跨：失败", u"一天：跨：all",
                      u"一天：跨：成功", u"一天：跨：失败", u"一周：all：all", u"一周：all：成功", u"一周：all：失败", u"一周：不跨：all", u"一周：不跨：成功", u"一周：不跨：失败",
                      u"一周：跨：all", u"一周：跨：成功", u"一周：跨：失败", u"一月：all：all", u"一月：all：成功", u"一月：all：失败", u"一月：不跨：all", u"一月：不跨：成功",
                      u"一月：不跨：失败", u"一月：跨：all", u"一月：跨：成功", u"一月：跨：失败"]
        head = ["ISI", "Name", "NE Name/Organization Name", "Organization Node", "Registration", "Deregistration"]
        rowtitle = 1
        rowhead = 1
        for i in range(0, len(title), 1):
            MSResgistrationCountSheet.write(rowtitle - 1, 0, title[i])
            rowtitle = rowtitle + len(head)
            for j in range(0, len(head), 1):
                MSResgistrationCountSheet.write(rowhead - 1, 1, head[j])
                rowhead = rowhead + 1
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 点击性能
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(1, column - 1)).click()  # 点击系统信息统计
        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(28, column - 1), 1).click()  # 点击系统终端登记次数
        nexttablerow = 1
        for dateindex in range(0, 3, 1):
            for i in range(0, len(self.InterSystemCombox), 1):
                for j in range(0, len(self.CallResultCombox), 1):
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i', 0.5).click()  # 高级搜索
                    time.sleep(0.5)
                    # 跳转到iframe ID变化的iframe
                    iframes = driver.find_elements_by_tag_name('iframe')
                    for iframe in iframes:
                        if 'layui-layer-iframe' in iframe.get_property('id'):
                            self.Func.switch_to_frame(driver, 'mainFrame,' + iframe.get_property('id'))
                    self.Func.ChooseDate(driver, self.StartDateArray[dateindex],'//*[@id="condition"]/table[1]/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                    self.Func.ChooseDate(driver, self.todaydate, '//*[@id="condition"]/table[1]/tbody/tr[4]/td/div/span/span/span')  # 截止日期
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[8]/td/div/span/span/span', '//*[@id="_easyui_combobox_i1_' + str(i) + '"]')  # 是否跨系统
                    self.Func.Select(driver, '//*[@id="condition"]/table/tbody/tr[10]/td/div/span/span/span', '//*[@id="_easyui_combobox_i2_' + str(j) + '"]')  # 呼叫结果
                    if self.objparam.cell_value(29, column - 1) == '//*[@id="checkSta"]':
                        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(29, column - 1)).click()  # 基站全选
                    elif self.objparam.cell_value(29, column - 1) == '//*[@id="ztree_1_check"]':
                        self.Func.find_element_by_xpath(driver, self.objparam.cell_value(29, column - 1), 1).click()  # 组织架构全选
                    self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确认
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    if self.Func.isElementExist(driver, 'layui-layer-btn0'):
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                    else:
                        trs = driver.find_elements_by_tag_name('tr')
                        rowcount = len(trs) - 1
                        firstcolumn = 3
                        for NMrow in range(0, rowcount, 1):
                            firstrow = nexttablerow
                            if column == 2:
                                MSResgistrationCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                                MSResgistrationCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                                MSResgistrationCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                                MSResgistrationCountSheet.write(firstrow + 4, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 5))
                            elif column == 3:
                                MSResgistrationCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                                MSResgistrationCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                                MSResgistrationCountSheet.write(firstrow + 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                                MSResgistrationCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                                MSResgistrationCountSheet.write(firstrow + 4, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 6))
                            elif column == 4:
                                MSResgistrationCountSheet.write(firstrow - 1, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,2))
                                MSResgistrationCountSheet.write(firstrow, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,3))
                                MSResgistrationCountSheet.write(firstrow + 1, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,4))
                                MSResgistrationCountSheet.write(firstrow + 2, firstcolumn - 1, self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,5))
                                MSResgistrationCountSheet.write(firstrow + 3, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]', NMrow + 1,6))
                                MSResgistrationCountSheet.write(firstrow + 4, firstcolumn - 1,self.Func.getTableCell(driver, '//*[@id="tb"]',NMrow + 1, 7))
                            firstcolumn = firstcolumn + 1
                    nexttablerow = nexttablerow + len(head)
        resxlsbook.save(os.getcwd() + "\\" + self.resfiles[column - 2] + ".xls")

    '''话务定制统计'''

    def Customizable_Statistics(self):
        # self.objparam = self.Func.read_excel('config.xls', '话务定制_系统基站组织架构信息统计配置')
        # 初始化数据
        item_name = ['Individual_GroupCallDuration', 'Intra_Inter-BSCallDuration', 'Intra_Inter-systemCallDuration',
                     'CallDurationbyCallType', 'Individual_GroupCallCount', 'Intra_Inter-BSCallCount',
                     'Intra_Inter-systemCallCount', 'CallCountbyCallType', ' Individual_GroupDataServiceTraffic',
                     'DataServiceTrafficbyCallType', 'Individual_GroupDataServiceCount', 'DataServiceCountbyCallType',
                     'TotalRegistrationCount', 'Registration_DeregistrationCount']
        item_number = [1, 2, 3, 4, 5, 6, 7, 8, 13, 16, 17, 20, 25, 27]  # 查找项对应的编号
        sheet_name = ['Call Duration', 'Call Count', 'Data Service Traffic', 'Data Service Count',
                      'Total Registration Count']
        Statistic_Object = ['System', 'BS', 'Organization']
        Statistic_query = ['mso', 'tsc', 'vpn']
        cycle = ['day', 'week', 'month']
        type_name_call = ['Group Call,', 'Individual Call,', 'Inter-BS,', 'Intra-BS,', 'Intra-system,', 'Inter-system,',
                          'Normal Call', 'GPS Group Call', 'Include Call', 'Patch Group Call', 'PSTN Call',
                          'Reserved Call', 'Broadcast Call', 'Full-duplex Call', 'Emergency Call', 'Encrypted Call',
                          'Ambience Listening']
        type_name_dataservice = ['Group Call', 'Individual Call', 'Individual Short Message Call',
                                 'Individual Packet Data Call', 'Individual Status Message Call',
                                 'Group Short Message Call', 'Group Packet Data Call', 'Group Status Message Call']
        type_name_registration = ['Local Registration,', 'Local Deregistration,', 'Inter-system Registration,',
                                  'Inter-system Deregistration,', 'Registration,', 'Deregistration,', ]
        period_name = ['呼叫类别', '最近一天', '最近一周', '最近一月']
        period_name_regestration = ['呼叫类别', '一天：all', '一天：成功', '一天：失败,', '一周：all', '一周：成功', '一周：失败', '一月：all', '一月：成功',
                                    '一月：失败']
        # 新建excel及初始化
        for i in range(len(Statistic_Object)):
            # 新建三个定制统计excel
            self.Func.new_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name)  # 新建三个定制统计excel
            # sheet页初始化
            for j in range(len(sheet_name)):
                # 呼叫类sheet页初始化
                if j == 0 or j == 1:  # 呼叫类sheet页初始化
                    if j == 0:
                        for n in range(len(period_name)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, n,
                                                  unicode(period_name[n]))  # 写行
                    else:
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, 0,
                                              unicode(period_name[0]))  # 写行
                        for n in range(1, len(period_name)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, n,
                                                  unicode(period_name[n]) + unicode('(次)'))  # 写行
                    for m in range(len(type_name_call)):
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], m + 1, 0,
                                              unicode(type_name_call[m]))  # 写列
                # 数据类sheet页初始化
                elif j == 2 or j == 3:  # 数据类heet页初始化
                    if j == 2:
                        for n in range(len(period_name)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, n,
                                                  unicode(period_name[n]))  # 写行
                    else:
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, 0,
                                              unicode(period_name[0]))  # 写行
                        for n in range(1, len(period_name)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, n,
                                                  unicode(period_name[n]) + unicode('(KB)'))  # 写行
                    for m in range(len(type_name_dataservice)):
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], m + 1, 0,
                                              unicode(type_name_dataservice[m]))  # 写列
                # 登记类sheet页初始化
                else:  # 登记类sheet页初始化
                    for n in range(len(period_name_regestration)):
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], 0, n,
                                              unicode(period_name_regestration[n]))
                    for m in range(len(type_name_registration)):
                        self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', sheet_name[j], m + 1, 0,
                                              unicode(type_name_registration[m]))
        # 读取配置文件数据
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))
        # print loginPageAddr
        username = str(configsheet.cell_value(1, 1))
        # print username
        password = str(configsheet.cell_value(1, 2))
        # print password

        configsheet1 = self.Func.read_excel('config.xls', '话务定制_系统基站组织架构信息统计配置')
        systemid = str(configsheet1.cell_value(45, 1))
        # print systemid
        Custo_Url = str(configsheet1.cell_value(46, 1))
        # 打开chrome
        driver = self.Func.OpenChrome(loginPageAddr)
        # 登录网管进入性能
        self.Func.login(driver, username, password)
        time.sleep(10)
        # 接口查询系统-基站-组织架构
        for i in range(len(Statistic_Object)):
            # 各种项类型查询
            for m in range(len(item_name)):
                # 呼叫类和数据类
                if m < 12:  # 呼叫类和数据类
                    item_data = {}
                    # 呼叫和数据类接口获取数据
                    for j in range(len(cycle)):  # 天、周、月
                        query_data = {'item': str(item_number[m]),
                                      'msoId': systemid,
                                      '_systemId': systemid,
                                      'type': Statistic_query[i],
                                      'createDate': str(
                                          time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))),
                                      'userName': username,
                                      'cycle': cycle[j],
                                      'id': '2', 'limitValue': '', 'msgCode': '', 'regResult': '', 'sign': '1',
                                      'subIds': '', 'viewKey': 'view1'
                                      }
                        data_name = Statistic_Object[i] + ' ' + item_name[m] + ' ' + cycle[j]
                        print Custo_Url
                        print query_data
                        data = self.Func.queryChartData(driver, Custo_Url, query_data)
                        item_data[cycle[
                            j]] = data  # [[u'2', u'18'],[u'126', u'205'],[u'158', u'364']]    【天【单呼，组呼】，周【单呼，组呼】，月【单呼，组呼】】
                    # 单组呼时长写入excel
                    if m == 0:
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[0])
                        Inidi_row_num = ParamsSheet.col_values(0).index('Individual Call,')  # 获取单呼行号
                        Group_row_num = ParamsSheet.col_values(0).index('Group Call,')  # 获取组呼行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  Inidi_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][0]))  # 单呼时长数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  Group_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][1]))  # 组呼时长数据写入
                    # 单跨站时长写入excel
                    elif m == 1:
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[0])
                        IntraBS_row_num = ParamsSheet.col_values(0).index('Intra-BS,')  # 获取单站行号
                        InterBS_row_num = ParamsSheet.col_values(0).index('Inter-BS,')  # 获取跨站行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  IntraBS_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][0]))  # 单站时长数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  InterBS_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][1]))  # 跨站时长数据写入
                    # 单跨系统时长写入excel
                    elif m == 2:
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[0])
                        IntraSYS_row_num = ParamsSheet.col_values(0).index('Intra-system,')  # 获取单站行号
                        InterSYS_row_num = ParamsSheet.col_values(0).index('Inter-system,')  # 获取跨站行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  IntraSYS_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][0]))  # 单系统数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                  InterSYS_row_num, j + 1,
                                                  self.Func.second2time(item_data[cycle[j]][1]))  # 跨系统数据写入
                    # 不同类型呼叫时长写入excel
                    elif m == 3:
                        Call_Type = ['Normal Call', 'Emergency Call', 'Encrypted Call', 'Full-duplex Call',
                                     'Broadcast Call', 'PSTN Call', 'Ambience Listening', 'Include Call',
                                     'GPS Group Call', 'Patch Group Call', 'Reserved Call']
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[0])
                        for calltype in range(len(Call_Type)):
                            print calltype
                            CallType_row_num = ParamsSheet.col_values(0).index(Call_Type[calltype])  # 获取每个类型行号
                            for j in range(len(cycle)):
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Duration',
                                                      CallType_row_num, j + 1,
                                                      self.Func.second2time(item_data[cycle[j]][calltype]))  # 不同类型时长写入
                    # 单组呼次数excel写入excel
                    elif m == 4:  # 单组呼次数
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[1])
                        Inidi_row_num = ParamsSheet.col_values(0).index('Individual Call,')  # 获取单呼行号
                        Group_row_num = ParamsSheet.col_values(0).index('Group Call,')  # 获取组呼行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  Inidi_row_num, j + 1, item_data[cycle[j]][0])  # 单呼次数数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  Group_row_num, j + 1, item_data[cycle[j]][1])  # 组呼次数数据写入
                    # 单跨站次数写入excel
                    elif m == 5:  # 单跨站次数
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[1])
                        IntraBS_row_num = ParamsSheet.col_values(0).index('Intra-BS,')  # 获取单站行号
                        InterBS_row_num = ParamsSheet.col_values(0).index('Inter-BS,')  # 获取跨站行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  IntraBS_row_num, j + 1, item_data[cycle[j]][0])  # 单站次数数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  InterBS_row_num, j + 1, item_data[cycle[j]][1])  # 跨站次数数据写入
                    # 单跨系统时长写入excel
                    elif m == 6:  # 单跨系统时长
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[1])
                        IntraSYS_row_num = ParamsSheet.col_values(0).index('Intra-system,')  # 获取单系统行号
                        InterSYS_row_num = ParamsSheet.col_values(0).index('Inter-system,')  # 获取跨系统行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  IntraSYS_row_num, j + 1, item_data[cycle[j]][0])  # 单系统次数数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                  InterSYS_row_num, j + 1, item_data[cycle[j]][1])  # 跨系统次数数据写入
                    # 不同类型呼叫次数写入excel
                    elif m == 7:  # 不同类型呼叫次数
                        Call_Type = ['Normal Call', 'Emergency Call', 'Encrypted Call', 'Full-duplex Call',
                                     'Broadcast Call', 'PSTN Call', 'Ambience Listening', 'Include Call',
                                     'GPS Group Call', 'Patch Group Call', 'Reserved Call']
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[1])
                        for calltype in range(len(Call_Type)):
                            CallType_row_num = ParamsSheet.col_values(0).index(Call_Type[calltype])  # 获取每个类型行号
                            for j in range(len(cycle)):
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Call Count',
                                                      CallType_row_num, j + 1,
                                                      item_data[cycle[j]][calltype])  # 不同类型次数数据写入
                    # 单组呼数据业务流量写入excel
                    elif m == 8:  # 单组呼数据业务流量
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[2])
                        Inidi_row_num = ParamsSheet.col_values(0).index('Individual Call')  # 获取单呼行号
                        Group_row_num = ParamsSheet.col_values(0).index('Group Call')  # 获取组呼行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                  'Data Service Traffic', Inidi_row_num, j + 1,
                                                  item_data[cycle[j]][0])  # 单呼次数数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                  'Data Service Traffic', Group_row_num, j + 1,
                                                  item_data[cycle[j]][1])  # 组呼次数数据写入
                    # 不同类型数据业务流量写入excel
                    elif m == 9:  # 不同类型数据业务流量
                        Call_Type = ['Individual Short Message Call', 'Individual Packet Data Call',
                                     'Individual Status Message Call', 'Group Short Message Call',
                                     'Group Packet Data Call', 'Group Status Message Call']
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[2])
                        for calltype in range(len(Call_Type)):
                            CallType_row_num = ParamsSheet.col_values(0).index(Call_Type[calltype])  # 获取每个类型行号
                            for j in range(len(cycle)):
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                      'Data Service Traffic', CallType_row_num, j + 1,
                                                      item_data[cycle[j]][calltype])  # 不同类型数据流量数据写入
                    # 单组呼数据业务次数写入excel
                    elif m == 10:  # 单组呼数据业务次数
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[3])
                        Inidi_row_num = ParamsSheet.col_values(0).index('Individual Call')  # 获取单呼行号
                        Group_row_num = ParamsSheet.col_values(0).index('Group Call')  # 获取组呼行号
                        for j in range(len(cycle)):
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Data Service Count',
                                                  Inidi_row_num, j + 1, item_data[cycle[j]][0])  # 单呼次数数据写入
                            self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls', 'Data Service Count',
                                                  Group_row_num, j + 1, item_data[cycle[j]][1])  # 组呼次数数据写入
                    # 不同类型数据业务次数写入excel
                    elif m == 11:  # 不同类型数据业务次数
                        Call_Type = ['Individual Short Message Call', 'Individual Packet Data Call',
                                     'Individual Status Message Call', 'Group Short Message Call',
                                     'Group Packet Data Call', 'Group Status Message Call']
                        ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                           sheet_name[3])
                        for calltype in range(len(Call_Type)):
                            CallType_row_num = ParamsSheet.col_values(0).index(Call_Type[calltype])  # 获取每个类型行号
                            for j in range(len(cycle)):
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                      'Data Service Count', CallType_row_num, j + 1,
                                                      item_data[cycle[j]][calltype])  # 不同类型数据次数数据写入


                                #
                                #
                                #
                                #
                                #
                # 登记类
                else:  # 登记类
                    result_num = ['4294967295', '0', '1']
                    item_data = {}
                    for result in range(len(result_num)):
                        # 登记类接口获取数据
                        for j in range(len(cycle)):  # 天、周、月
                            query_data = {'item': str(item_number[m]),
                                          'msoId': systemid,
                                          '_systemId': systemid,
                                          'type': Statistic_query[i],
                                          'createDate': str(
                                              time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))),
                                          'userName': username,
                                          'cycle': cycle[j],
                                          'regResult': result_num[result],
                                          'id': '2', 'limitValue': '', 'msgCode': '', 'sign': '1',
                                          'subIds': '', 'viewKey': 'view1'
                                          }
                            data_name = Statistic_Object[i] + ' ' + item_name[m] + ' ' + cycle[j] + result_num[result]
                            data = self.Func.queryChartData(driver, Custo_Url, query_data)
                            item_data[cycle[j]] = data
                        # 登记总次数
                        # 登记总次数写入excel
                        if m == 12:
                            Call_Type = ['Local Registration,', 'Local Deregistration,', 'Inter-system Registration,',
                                         'Inter-system Deregistration,']
                            ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                               sheet_name[4])
                            for calltype in range(len(Call_Type)):
                                CallType_row_num = ParamsSheet.col_values(0).index(Call_Type[calltype])
                                print CallType_row_num  # 获取每个类型行号
                                for jj in range(len(cycle)):
                                    self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                          'Total Registration Count', CallType_row_num,
                                                          3 * jj + result + 1,
                                                          item_data[cycle[jj]][calltype])  # 不同登记类型次数数据写入
                        # 登记/去登记次数写入excel
                        else:
                            ParamsSheet = self.Func.read_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                               sheet_name[4])
                            Re_row_num = ParamsSheet.col_values(0).index('Registration,')  # 获取登记行号
                            DeRe_row_num = ParamsSheet.col_values(0).index('Deregistration,')  # 获取去登记行号
                            for j in range(len(cycle)):
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                      'Total Registration Count', Re_row_num, 3 * j + result + 1,
                                                      item_data[cycle[j]][0])  # 登记次数数据写入
                                self.Func.write_excel('Customizable_' + Statistic_Object[i] + '.xls',
                                                      'Total Registration Count', DeRe_row_num, 3 * j + result + 1,
                                                      item_data[cycle[j]][1])  # 去登记次数数据写入
        time.sleep(2)
        driver.quit()

    # def KPI_Statistic(self):
    #     configsheet = self.Func.read_excel('config.xls', 'Sheet1')
    #     loginPageAddr = str(configsheet.cell_value(1, 0))
    #     username = str(configsheet.cell_value(1, 1))
    #     password = str(configsheet.cell_value(1, 2))
    #     # 打开chrome
    #     driver = self.Func.OpenChrome(loginPageAddr)
    #     # 登录网管进入性能
    #     self.Func.loginAndtoEOES(driver, username, password)
    '''KPI统计'''

    def KPI_Statistic(self):

        KPI_COL = ['Individual Call Duration', 'Individual Call Count', 'Group Call Duration', 'Group Call Count',
                   'Individual Data Service Traffic (KB)', 'Individual Data Service Count',
                   'Group Data Service Traffic (KB)', 'Group Data Service Count', 'Successful Registration Count',
                   'Failed Registration Count', 'Successful Deregistration Count', 'Failed Deregistration Count',
                   'Inter-BS Call Duration', 'Inter-BS Call Count', 'Inter-system Call Duration',
                   'Inter-system Call Count', 'Intra-BS Call Duration', 'Intra-BS Call Count',
                   'Intra-system Call Duration', 'Intra-system Call Count']
        KPI_BS_Row = ['NE Name']
        KPI_ORA = ['CallType']
        Statistic_Object = ['System', 'BS', 'Organization']
        Sheet_Name = ['Day', 'Week', 'Month']
        cycle = ['day', 'week', 'month']

        lastmonth = self.Func.getLastMonth()
        lastweek = self.Func.getLastWeek()
        # 新建excel并初始化
        for i in range(len(Statistic_Object)):
            self.Func.new_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name)
            for j in range(len(Sheet_Name)):

                # 初始化写入excle
                if i == 1:
                    self.Func.write_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name[j], 0, 0, KPI_BS_Row)
                    for m in range(len(KPI_COL)):
                        self.Func.write_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name[j], m + 1, 0,
                                              KPI_COL[m])
                elif i == 2:
                    self.Func.write_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name[j], 0, 0, KPI_ORA)
                    for m in range(len(KPI_COL)):
                        self.Func.write_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name[j], m + 1, 0,
                                              KPI_COL[m])
                else:
                    for m in range(len(KPI_COL)):
                        self.Func.write_excel('KPI_' + Statistic_Object[i] + '.xls', Sheet_Name[j], m, 0, KPI_COL[m])

        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))
        username = str(configsheet.cell_value(1, 1))
        password = str(configsheet.cell_value(1, 2))

        for obj_num in range(len(Statistic_Object)):
            for cycle_num in range(len(cycle)):
                # for cycle_num in range(2,3):
                self.Func.CloseProcess('NMC,EXCEL,chromedriver,chrome')
                # 打开chrome
                driver = self.Func.OpenChrome(loginPageAddr)
                # 登录网管进入性能
                self.Func.login(driver, username, password)
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 进入性能
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[7]').click()  # KPI
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[7]/div/div[' + str(
                    obj_num + 1) + ']').click()  # 系统KPI
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[2]/div[1]/i').click()  # 高级搜索
                # self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.switch_to_frame(driver, 'mainFrame,layui-layer-iframe1')

                if obj_num == 1:  # 如果基站KPI则全选所有基站
                    self.Func.find_element_by_xpath(driver, '//*[@id="checkSta"]').click()  # 基站全选
                elif obj_num == 2:  # 如果组织架构KPI则全选组织架构节点
                    self.Func.find_element_by_xpath(driver, '//*[@id="ztree_1_check"]').click()  # 组织架构全选
                else:
                    pass
                if cycle_num == 1:
                    self.Func.ChooseDate(driver, lastweek,
                                         '//*[@id="condition"]/table/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                elif cycle_num == 2:
                    self.Func.ChooseDate(driver, lastmonth,
                                         '//*[@id="condition"]/table/tbody/tr[2]/td/div/span/span/span')  # 开始日期
                else:
                    pass

                self.Func.find_element_by_xpath(driver, '//*[@id="main_div"]/div[2]/div/a[1]').click()  # 确定
                self.Func.switch_to_frame(driver, 'mainFrame')
                table = driver.find_element_by_id('dataTable')
                table_cols = table.find_elements_by_tag_name('th')
                table_rows = driver.find_element_by_id('tb').find_elements_by_tag_name('tr')
                print "总行数:", len(table_rows)
                print "总列数：", len(table_cols)
                KPI_Statistic = []
                # 获取界面表格数据，并按顺序存储在KPI_Statistic
                for i in range(1, len(table_rows) + 1):  # 行
                    KPIdata = []
                    data = []
                    for j in range(2, len(table_cols) + 1):  # 列数
                        header = self.Func.find_element_by_xpath(driver, '//*[@id="dataTable"]/thead/tr/th[' + str(
                            j) + ']').text.decode("utf-8")  # 获取表头
                        print i, '   ', j
                        if i == 1:
                            data.append(self.Func.find_element_by_xpath(driver,
                                                                        '//*[@id="dataTable"]/tbody/tr/td[' + str(
                                                                            j) + ']').text.decode("utf-8"))  # 获取单元格数据
                        else:
                            data.append(self.Func.find_element_by_xpath(driver, '//*[@id="dataTable"]/tbody/tr[' + str(
                                i) + ']/td[' + str(j) + ']').text.decode("utf-8"))  # 获取单元格数据
                    # KPIdata.append(data)
                    # print KPIdata
                    KPI_Statistic.append(data)  # 整个表格 不包含表头
                    print data  # 一行的数据
                for row in range(len(KPI_Statistic)):  # UI中的行
                    for col in range(len(KPI_Statistic[0])):  # UI中的列
                        self.Func.write_excel('KPI_' + Statistic_Object[obj_num] + '.xls', Sheet_Name[cycle_num], col,
                                              row + 1, KPI_Statistic[row][col])


    '''业务详单统计'''
    def DetailedListStatistics(self):
        # 从配置文件获取参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = configsheet.cell_value(1,0)  # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        fileaddr = configsheet.cell_value(4, 5)  # 导出地址
        driver = self.Func.OpenChrome(loginPageAddr, fileaddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        # 删除文件夹下之前生成的数据文件
        self.Func.del_files(fileaddr, u'语音')
        self.Func.del_files(fileaddr, u'数据')
        '''语音业务'''
        StartDate = configsheet.cell_value(4,1)  # 开始日期
        StartTime = configsheet.cell_value(4,2)  # 开始时间
        EndDate = configsheet.cell_value(4,3)  # 结束日期
        EndTime = configsheet.cell_value(4,4)  # 结束时间
        RunNum = configsheet.cell_value(4, 6)  # 执行次数
        #  比较字段
        CallType = configsheet.cell_value(8, 1)  # 呼叫类型
        CallModel = configsheet.cell_value(8, 2)  # 呼叫模式
        CallerNum = configsheet.cell_value(8, 3)  # 主叫号码
        CallerType = configsheet.cell_value(8, 4)  # 主叫用户类型
        CalledNum = configsheet.cell_value(8, 5)  # 被叫号码
        CalledType = configsheet.cell_value(8, 6)  # 被叫用户类型
        CallResult = configsheet.cell_value(8, 7)  # 呼叫结果
        TCLjudgedata = configsheet.cell_value(10, 1)  # 判断数据
        CaseRunXls = configsheet.cell_value(4, 7)  # 用例Excel文件名
        TCLResultXls = configsheet.cell_value(4, 8)  # TCL日志Excel文件名
        NmsResultXls = configsheet.cell_value(4, 9)  # 网管导出的语音业务的数据合并后文件名
        CompareResultXls = configsheet.cell_value(4, 10)  # Tcl和网管比较结果文件名
        NmsCompareResultXls = configsheet.cell_value(4, 11)  # 网管语音业务的数据和导出数据比较结果文件名
        print u'语音 开始日期：'+StartDate,StartTime + u' \n     截止日期：'+EndDate,EndTime
        #　查询数据
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 性能
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[5]/span').click()  # 业务详单统计
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[5]/div/div[1]').click()  #语音业务详单
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[1]/div[1]/div[1]/i').click()  # 高级搜索
        self.Func.switch_to_frame(driver, 'mainFrame,layui-layer-iframe1')
        self.Func.ChooseDate(driver, StartDate, '//*[@id="condition"]/table/tbody/tr[2]/td/div/span[1]/span/span') # 开始日期
        self.Func.ChooseDate(driver, EndDate, '//*[@id="condition"]/table/tbody/tr[4]/td/div/span[1]/span/span') # 结束日期
        self.Func.ChooseTime(driver, StartTime, '//*[@id="condition"]/table/tbody/tr[2]/td/div/span[2]/span/span', '//*[@id="_easyui_combobox_i1_')  # 开始时间
        self.Func.ChooseTime(driver, EndTime, '//*[@id="condition"]/table/tbody/tr[4]/td/div/span[2]/span/span', '//*[@id="_easyui_combobox_i2_') # 结束时间
        self.Func.find_element_by_xpath(driver, '//*[@id="checkSta"]').click() # 基站全选
        self.Func.find_element_by_xpath(driver, '//*[@id="commit"]').click() # 确认键
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.select_by_visible_text(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[1]/select', '3000')  # 分页 每页显示最多3000数据

        '''语音业务导出'''
        DayDiff = self.Func.getDayDiff(StartDate, EndDate)
        print u"时间间隔天数：" + str(DayDiff)
        files = []
        index = 0
        nmscompareresxlsrows = 1
        nmscompareresxlsbook = xlwt.Workbook()
        nmscompareresxlssheet = nmscompareresxlsbook.add_sheet('Sheet1')
        nmscompareresxlssheet.write(0,0,u'文件名')
        nmscompareresxlssheet.write(0,1,CalledNum)
        nmscompareresxlssheet.write(0,2,u'比较结果')

        for singleday in range(1,DayDiff + 1,1):
            print u"第" + str(singleday) +u"天：",
            # 分割总数据及显示的数据量
            text = self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/div[1]').text  # 显示 1 到 38 共 38 条
            totaldata = text.split()[5]  # 38
            if totaldata == 0:
                pass
            else:
                # 获取总页数
                Pages = int(self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[7]/span').__getattribute__("text")[1:])
                print u"(共" + str(Pages) + u"页)"
                for singlepage in range(Pages,0,-1):
                    print u'  第' + str(singlepage) + u'页 ',
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').clear()
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').send_keys(singlepage)  # 跳转到第1页
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').send_keys(Keys.ENTER)
                    self.Func.find_element_by_xpath(driver, '//*[@id="checkAll"]').click()  # 全选当前页表格数据
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[1]/div[1]/div[2]', 3).click()  # 点击导出
                    text1 = self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/div[1]').text  # 显示 1 到 38 共 38 条
                    databegin = text1.split()[1]  # 1
                    dataend = text1.split()[3]  # 38
                    currentpagenum = int(dataend) - int(databegin) + 1
                    time.sleep(3)
                    try:
                        shutil.move(fileaddr + u"\语音业务详单.xlsx",fileaddr + u"\语音_" + str(datetime.strptime(StartDate, "%Y-%m-%d") + timedelta(days=(singleday - 1))).replace("-","_").split()[0] + "_Page" + str(singlepage) + ".xlsx")
                    except:
                        shutil.move(fileaddr + u"\Voice Service.xlsx", fileaddr + u"\语音_" + str(datetime.strptime(StartDate, "%Y-%m-%d") + timedelta(days=(singleday - 1))).replace("-","_").split()[0] + "_Page" + str(singlepage) + ".xlsx")
                    files.append(fileaddr + u"\语音_" + str(datetime.strptime(StartDate, "%Y-%m-%d") + timedelta(days=(singleday - 1))).replace("-","_").split()[0] + "_Page" + str(singlepage) + ".xlsx")

                    '''比较网管数据和导出数据是否一致'''
                    singlepagexlsbook = xlrd.open_workbook(files[index])
                    singlepagexlssheet = singlepagexlsbook.sheet_by_index(0)
                    singlepagexlscolumns = singlepagexlssheet.ncols
                    singlepagexlsrows = singlepagexlssheet.nrows
                    print u'网管导出数据行数：'+str(singlepagexlsrows - 1)
                    NmsDataNum = currentpagenum
                    RowNum,callednumColNum = self.Func.GetCellRowAndCol(singlepagexlssheet, CalledNum, 0, singlepagexlscolumns)
                    singlepagexlsrowsIndex = 1
                    print u"    正在比较（行）：",
                    if singlepagexlsrows == NmsDataNum + 1:
                        while singlepagexlsrowsIndex < singlepagexlsrows:
                            print singlepagexlsrowsIndex,
                            singlepagexlscolumnsIndex = 0
                            while singlepagexlscolumnsIndex < singlepagexlscolumns:
                                nmsdata = self.Func.getTableCell(driver, '//*[@id="tb"]', singlepagexlsrowsIndex, singlepagexlscolumnsIndex + 2)
                                if singlepagexlssheet.cell_value(singlepagexlsrowsIndex, singlepagexlscolumnsIndex).strip() == nmsdata:
                                    singlepagexlscolumnsIndex = singlepagexlscolumnsIndex + 1
                                else:
                                    nmscompareresxlssheet.write(nmscompareresxlsrows, 0, files[index])
                                    nmscompareresxlssheet.write(nmscompareresxlsrows, 1, singlepagexlssheet.cell_value(singlepagexlsrowsIndex, callednumColNum))
                                    nmscompareresxlssheet.write(nmscompareresxlsrows, 2, u"数据不一致")
                                    nmscompareresxlsrows = nmscompareresxlsrows + 1
                                    singlepagexlscolumnsIndex = singlepagexlscolumnsIndex + 1
                                    break
                            if singlepagexlscolumnsIndex == singlepagexlscolumns:
                                nmscompareresxlssheet.write(nmscompareresxlsrows, 0, files[index])
                                nmscompareresxlssheet.write(nmscompareresxlsrows, 1, singlepagexlssheet.cell_value(singlepagexlsrowsIndex, callednumColNum))
                                nmscompareresxlssheet.write(nmscompareresxlsrows, 2, u"数据一致")
                                nmscompareresxlsrows = nmscompareresxlsrows + 1
                                singlepagexlscolumnsIndex = singlepagexlscolumnsIndex + 1
                            singlepagexlsrowsIndex = singlepagexlsrowsIndex + 1
                    else:
                        nmscompareresxlssheet.write(nmscompareresxlsrows, 0, files[index])
                        nmscompareresxlssheet.write(nmscompareresxlsrows, 2, u"网管当前页数据和导出数据的总数不一致")
                        nmscompareresxlsrows = nmscompareresxlsrows + 1
                    print '.'
                    index = index + 1
            if singleday <> DayDiff:
                self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[11]/a/span/span[2]').click() # 点击下一天
        self.Func.del_files(fileaddr, NmsCompareResultXls)
        nmscompareresxlsbook.save(str(fileaddr) +'\\'+ str(NmsCompareResultXls))

        '''将网管导出的语音业务数据合并到一个Excel'''
        resxlsbook = xlwt.Workbook()
        resxlssheet = resxlsbook.add_sheet('Sheet1')

        startrow = 0

        print u"文件数：" + str(len(files))

        for file in range(1,len(files) + 1, 1):
            print u"   正在合并文件："+ str(files[file - 1])
            xlsbook = xlrd.open_workbook(files[file - 1])
            xlssheet = xlsbook.sheet_by_index(0)
            rows = xlssheet.nrows
            columns = xlssheet.ncols

            if file == 1:
                for j in range(0, columns, 1):
                    resxlssheet.write(0, j, xlssheet.cell_value(0, j))
            for i in range(1, rows, 1):
                for j in range(0, columns, 1):
                    resxlssheet.write(startrow + i, j, xlssheet.cell_value(rows - i, j))
            startrow = startrow + rows - 1
        print u"合并文件完成。"
        self.Func.del_files(fileaddr, NmsResultXls)
        resxlsbook.save(fileaddr + "\\" + NmsResultXls)

        '''数据业务'''
        D_StartDate = configsheet.cell_value(5,1)
        D_StartTime = configsheet.cell_value(5,2)
        D_EndDate = configsheet.cell_value(5,3)
        D_EndTime = configsheet.cell_value(5,4)
        D_fileaddr = configsheet.cell_value(5,5)
        D_RunNum = configsheet.cell_value(5,6)
        D_CallType = configsheet.cell_value(9,1)
        D_CallerNum = configsheet.cell_value(9,2)
        D_CallerType = configsheet.cell_value(9,3)
        D_CalledNum = configsheet.cell_value(9,4)
        D_CalledType = configsheet.cell_value(9,5)
        D_SendResult = configsheet.cell_value(9,6)
        D_NmsResultXls = configsheet.cell_value(6,9)
        D_NmsCompareResultXls = configsheet.cell_value(6,11)
        print u'数据 开始日期：' + D_StartDate, D_StartTime + u' \n     截止日期：' + D_EndDate, D_EndTime

        '''数据业务导出'''
        driver.switch_to.default_content()
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]').click()  # 性能
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[5]/span').click()  # 业务详单统计
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[6]/div/div[5]/div/div[2]').click()  # 数据业务详单
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[1]/div[1]/div[1]/i').click()  # 高级搜索
        self.Func.switch_to_frame(driver, 'mainFrame,layui-layer-iframe1')
        self.Func.ChooseDate(driver, D_StartDate, '//*[@id="condition"]/table/tbody/tr[2]/td/div/span[1]/span/span')  # 开始日期
        self.Func.ChooseDate(driver, D_EndDate, '//*[@id="condition"]/table/tbody/tr[4]/td/div/span[1]/span/span')  # 结束日期
        self.Func.ChooseTime(driver, D_StartTime, '//*[@id="condition"]/table/tbody/tr[2]/td/div/span[2]/span/span', '//*[@id="_easyui_combobox_i1_')  # 开始时间
        self.Func.ChooseTime(driver, D_EndTime, '//*[@id="condition"]/table/tbody/tr[4]/td/div/span[2]/span/span', '//*[@id="_easyui_combobox_i2_')  # 结束时间
        self.Func.find_element_by_xpath(driver, '//*[@id="checkSta"]').click()  # 基站全选
        self.Func.find_element_by_xpath(driver, '//*[@id="commit"]').click()  # 确认键
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.select_by_visible_text(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[1]/select', '3000')  # 分页 每页显示最多3000数据

        '''数据业务导出'''
        D_DayDiff = self.Func.getDayDiff(D_StartDate, D_EndDate)
        print u"时间间隔天数：" + str(D_DayDiff)
        D_files = []
        D_index = 0
        D_nmscompareresxlsrows = 1

        D_nmscompareresxlsbook = xlwt.Workbook()
        D_nmscompareresxlssheet = D_nmscompareresxlsbook.add_sheet('Sheet1')
        D_nmscompareresxlssheet.write(0,0,u'文件名')
        D_nmscompareresxlssheet.write(0,1,D_CallType)
        D_nmscompareresxlssheet.write(0,2,u'比较结果')

        for D_singleday in range(1, D_DayDiff + 1, 1):
            print u"第" + str(D_singleday) +u"天：",
            # 分割总数据及显示的数据量
            D_text = self.Func.find_element_by_xpath(driver,'//*[@id="main"]/div[3]/div/div[1]').text  # 显示 1 到 38 共 38 条
            D_totaldata = D_text.split()[5]  # 38
            if D_totaldata == 0:
                pass
            else:
                # 获取总页数
                D_Pages = int(self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[7]/span').__getattribute__("text")[1:])
                print u"(共" + str(D_Pages) + u"页)"
                for D_singlepage in range(D_Pages,0,-1):
                    print u'  第' + str(D_singlepage) + u'页 ',
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').clear()
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').send_keys(D_singlepage)  # 跳转到第1页
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[6]/input').send_keys(Keys.ENTER)
                    self.Func.find_element_by_xpath(driver, '//*[@id="checkAll"]').click()  # 全选当前页表格数据
                    self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[1]/div[1]/div[2]', 3).click()  # 点击导出
                    D_text1 = self.Func.find_element_by_xpath(driver,'//*[@id="main"]/div[3]/div/div[1]').text  # 显示 1 到 38 共 38 条
                    D_databegin = D_text1.split()[1]  # 1
                    D_dataend = D_text1.split()[3]  # 38
                    D_currentpagenum = int(D_dataend) - int(D_databegin) + 1
                    time.sleep(3)
                    try:
                        shutil.move(D_fileaddr + u"\数据业务详单.xlsx",D_fileaddr + u"\数据_" + str(datetime.strptime(D_StartDate, "%Y-%m-%d") + timedelta(days=(D_singleday - 1))).replace("-","_").split()[0] + "_Page" + str(D_singlepage) + ".xlsx")
                    except:
                        shutil.move(D_fileaddr + u"\Data Service.xlsx", D_fileaddr + u"\数据_" + str(datetime.strptime(D_StartDate, "%Y-%m-%d") + timedelta(days=(D_singleday - 1))).replace("-","_").split()[0] + "_Page" + str(D_singlepage) + ".xlsx")
                    D_files.append(D_fileaddr + u"\数据_" + str(datetime.strptime(D_StartDate, "%Y-%m-%d") + timedelta(days=(D_singleday - 1))).replace("-","_").split()[0] + "_Page" + str(D_singlepage) + ".xlsx")

                    '''比较网管数据和导出数据是否一致'''
                    D_singlepagexlsbook = xlrd.open_workbook(D_files[D_index])
                    D_singlepagexlssheet = D_singlepagexlsbook.sheet_by_index(0)
                    D_singlepagexlscolumns = D_singlepagexlssheet.ncols
                    D_singlepagexlsrows = D_singlepagexlssheet.nrows
                    print u'网管导出数据行数：'+str(D_singlepagexlsrows - 1)
                    D_NmsDataNum = D_currentpagenum
                    D_RowNum,D_CallTypeColNum = self.Func.GetCellRowAndCol(D_singlepagexlssheet, D_CallType, 0, D_singlepagexlscolumns)
                    D_singlepagexlsrowsIndex = 1
                    print u"    正在比较（行）：",
                    if D_singlepagexlsrows == D_NmsDataNum + 1:
                        while D_singlepagexlsrowsIndex < D_singlepagexlsrows:
                            print D_singlepagexlsrowsIndex,
                            D_singlepagexlscolumnsIndex = 0
                            while D_singlepagexlscolumnsIndex < D_singlepagexlscolumns:
                                D_nmsdata = self.Func.getTableCell(driver, '//*[@id="tb"]', D_singlepagexlsrowsIndex + 1, D_singlepagexlscolumnsIndex + 2)
                                if D_singlepagexlssheet.cell_value(D_singlepagexlsrowsIndex, D_singlepagexlscolumnsIndex).strip() == D_nmsdata:
                                    D_singlepagexlscolumnsIndex = D_singlepagexlscolumnsIndex + 1
                                else:
                                    D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 0, D_files[D_index])
                                    D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 1, D_singlepagexlssheet.cell_value(D_singlepagexlsrowsIndex, D_CallTypeColNum))
                                    D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 2, u"数据不一致")
                                    D_nmscompareresxlsrows = D_nmscompareresxlsrows + 1
                                    D_singlepagexlscolumnsIndex = D_singlepagexlscolumnsIndex + 1
                                    break
                            if D_singlepagexlscolumnsIndex == D_singlepagexlscolumns:
                                D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 0, D_files[D_index])
                                D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 1, D_singlepagexlssheet.cell_value(D_singlepagexlsrowsIndex, D_CallTypeColNum))
                                D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 2, u"数据一致")
                                D_nmscompareresxlsrows = D_nmscompareresxlsrows + 1
                                D_singlepagexlscolumnsIndex = D_singlepagexlscolumnsIndex + 1
                            D_singlepagexlsrowsIndex = D_singlepagexlsrowsIndex + 1
                    else:
                        D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 0, D_files[D_index])
                        D_nmscompareresxlssheet.write(D_nmscompareresxlsrows, 2, u"网管当前页数据和导出数据的总数不一致")
                        D_nmscompareresxlsrows = D_nmscompareresxlsrows + 1
                    print '.'
                    D_index = D_index + 1
            if D_singleday <> D_DayDiff:
                self.Func.find_element_by_xpath(driver, '//*[@id="main"]/div[3]/div/table/tbody/tr/td[11]/a/span/span[2]').click() # 点击下一天
        self.Func.del_files(D_fileaddr, D_NmsCompareResultXls)
        D_nmscompareresxlsbook.save(str(D_fileaddr) +'\\'+ str(D_NmsCompareResultXls))

        '''将网管导出的语音业务数据合并到一个Excel'''
        D_resxlsbook = xlwt.Workbook()
        D_resxlssheet = D_resxlsbook.add_sheet('Sheet1')

        D_startrow = 0

        print u"文件数：" + str(len(D_files))

        for D_file in range(1,len(D_files) + 1, 1):
            print u"   正在合并文件："+ str(D_files[D_file - 1])
            D_xlsbook = xlrd.open_workbook(D_files[D_file - 1])
            D_xlssheet = D_xlsbook.sheet_by_index(0)
            D_rows = D_xlssheet.nrows
            D_columns = D_xlssheet.ncols

            if D_file == 1:
                for j in range(0, D_columns, 1):
                    D_resxlssheet.write(0, j, D_xlssheet.cell_value(0, j))
            for i in range(1, D_rows, 1):
                for j in range(0, D_columns, 1):
                    D_resxlssheet.write(D_startrow + i, j, D_xlssheet.cell_value(D_rows - i, j))
            D_startrow = D_startrow + D_rows - 1
        print u"合并文件完成。"
        self.Func.del_files(D_fileaddr, D_NmsResultXls)
        D_resxlsbook.save(D_fileaddr + "\\" + D_NmsResultXls)

        '''比较TCL和网管导出的数据'''
        NmsResultxlsbook = xlrd.open_workbook(fileaddr + "\\" + NmsResultXls)
        NmsResultxlssheet = NmsResultxlsbook.sheet_by_index(0)
        NmsResultxlscolumns = NmsResultxlssheet.ncols
        NmsResultxlsrows = NmsResultxlssheet.nrows

        D_NmsResultxlsbook = xlrd.open_workbook(D_fileaddr + "\\" + D_NmsResultXls)
        D_NmsResultxlssheet = D_NmsResultxlsbook.sheet_by_index(0)
        D_NmsResultxlscolumns = D_NmsResultxlssheet.ncols
        D_NmsResultxlsrows = D_NmsResultxlssheet.nrows

        RowNum, CallTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CallType, 0, NmsResultxlscolumns)
        RowNum, CallModelColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CallModel, 0, NmsResultxlscolumns)
        RowNum, CallerNumColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CallerNum, 0, NmsResultxlscolumns)
        RowNum, CallerTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CallerType, 0, NmsResultxlscolumns)
        RowNum, CalledNumColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CalledNum, 0, NmsResultxlscolumns)
        RowNum, CalledTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CalledType, 0, NmsResultxlscolumns)
        RowNum, CallResultColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, CallResult, 0, NmsResultxlscolumns)


        RowNum, D_CallTypeColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_CallType, 0, D_NmsResultxlscolumns)
        RowNum, D_CallerNumColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_CallerNum, 0, D_NmsResultxlscolumns)
        RowNum, D_CallerTypeColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_CallerType, 0, D_NmsResultxlscolumns)
        RowNum, D_CalledNumColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_CalledNum, 0, D_NmsResultxlscolumns)
        RowNum, D_CalledTypeColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_CalledType, 0, D_NmsResultxlscolumns)
        RowNum, D_SendResultColNum = self.Func.GetCellRowAndCol(D_NmsResultxlssheet, D_SendResult, 0, D_NmsResultxlscolumns)

        caserunxlsbook = xlrd.open_workbook(fileaddr + "\\" + CaseRunXls)
        caserunxlssheet = caserunxlsbook.sheet_by_name("Sheet1")
        CaseRunxlscolumns = caserunxlssheet.ncols
        CaseRunxlsrows = caserunxlssheet.nrows

        RowNum, caserunjudgedataColNum = self.Func.GetCellRowAndCol(caserunxlssheet, TCLjudgedata, 1, CaseRunxlscolumns)

        TCLResultxlsbook = xlrd.open_workbook(fileaddr + "\\" + TCLResultXls)
        TCLResultxlssheet = TCLResultxlsbook.sheet_by_index(0)
        TCLResultxlscolumns = TCLResultxlssheet.ncols
        TCLResultxlsrows = TCLResultxlssheet.nrows

        RowNum, TCLjudgedataColNum = self.Func.GetCellRowAndCol(TCLResultxlssheet, TCLjudgedata, 0, TCLResultxlscolumns)

        compareresxlsbook = xlwt.Workbook()
        compareresxlssheet = compareresxlsbook.add_sheet('Sheet1', cell_overwrite_ok=True)
        compareresxlssheet.write(0, 0, u"业务名")
        compareresxlssheet.write(0, 1, u"比较结果")

        calledRows = 3
        TCLResultRows = 2
        NmsresultRows = 2
        D_NmsresultRows = 2
        while calledRows < CaseRunxlsrows + 1:
            print CaseRunxlsrows
            print calledRows
            runnumIndex = 0
            runnumIndex1 = 0
            compareresxlssheet.write(calledRows - 2, 0, caserunxlssheet.cell_value(calledRows - 1, 1))
            if caserunxlssheet.cell_value(calledRows - 1, 1) == "登记":
                if TCLResultxlssheet.cell_value(TCLResultRows - 1, 2) == "登记":
                    compareresxlssheet.write(calledRows - 2, 1, TCLResultxlssheet.cell_value(TCLResultRows - 1,TCLjudgedataColNum - 1))
                    TCLResultRows = TCLResultRows + 1
            else:
                if TCLResultRows < TCLResultxlsrows + 1:
                    '''分割判断数据'''
                    dataarr = TCLResultxlssheet.cell_value(TCLResultRows - 1, TCLjudgedataColNum).split('：')
                    servicetype = dataarr[0]
                    judgedata = dataarr[1]
                    if servicetype == "语音":
                        if NmsresultRows < NmsResultxlsrows + 1:
                            JudgeDataDetail = judgedata.split('-')

                            '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                            for enlishi in range(0, len(JudgeDataDetail)):
                                JudgeDataDetail[enlishi] = JudgeDataDetail[enlishi].replace("=", "-")
                            ''''''''''''''''''''''''''''''''''''''''''''
                            print judgedata
                            CallTypeData = JudgeDataDetail[0]
                            CallModelData = JudgeDataDetail[1]
                            CallerNumData = JudgeDataDetail[2]
                            CallerTypeData = JudgeDataDetail[3]
                            CalledNumData = JudgeDataDetail[4]
                            CalledTypeData = JudgeDataDetail[5]
                            CallResultData = JudgeDataDetail[6]
                            if len(JudgeDataDetail) == 8:
                                WhetherServiceType = JudgeDataDetail[7]
                                WhetherServiceType_backup = WhetherServiceType
                                RowNum, WhetherServiceTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, WhetherServiceType, 0, NmsResultxlscolumns)
                            caseruncallednum = caserunxlssheet.cell_value(calledRows - 1, 3)
                            if caseruncallednum == "*1957":
                                caseruncallednum = "16777212"
                            elif caseruncallednum == "*1987" or caseruncallednum == "*******":
                                caseruncallednum = "16777215"
                            elif "*9*" in caseruncallednum or "*11*" in caseruncallednum:
                                arrnum = caseruncallednum.split('*')
                                caseruncallednum = arrnum[2]
                            elif "#9*" in caseruncallednum or "#1*" in caseruncallednum:
                                arrnum = caseruncallednum.split('*')
                                caseruncallednum = arrnum[1]
                            print caseruncallednum
                            print CalledNumData
                            CalledNumData_backup = CalledNumData

                            NmsIndex = 0
                            TclIndex = 0

                            while NmsresultRows < NmsResultxlsrows + 1:
                                if NmsresultRows + NmsIndex - 1 < NmsResultxlsrows:
                                    if caseruncallednum == NmsResultxlssheet.cell_value(NmsresultRows + NmsIndex - 1, CalledNumColNum).strip():
                                        NmsIndex = NmsIndex + 1
                                    else:
                                        break
                                else:
                                    break
                            while TCLResultRows < TCLResultxlsrows + 1:
                                if caseruncallednum == CalledNumData_backup:
                                    TclIndex = TclIndex + 1
                                    if TCLResultRows + TclIndex < TCLResultxlsrows + 1:
                                        dataarrTem = TCLResultxlssheet.cell_value(TCLResultRows + TclIndex - 1, TCLjudgedataColNum).split('：')
                                        servicetypeTem = dataarrTem[0]
                                        judgedataTem = dataarrTem[1]
                                        if servicetypeTem == "语音":
                                            #print "下一条数据:"
                                            #print judgedataTem
                                            JudgeDataDetailTem = judgedataTem.split('-')

                                            '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                            for enlishi in range(0, len(JudgeDataDetailTem)):
                                                JudgeDataDetailTem[enlishi] = JudgeDataDetailTem[enlishi].replace("=", "-")
                                            ''''''''''''''''''''''''''''''''''''''''''''

                                            CalledNumData_backup = JudgeDataDetailTem[4]
                                        else:
                                            break
                                else:
                                    break
                            print NmsIndex
                            print TclIndex
                            if NmsIndex <> TclIndex:
                                print "one"
                                NmsresultRows = NmsresultRows + NmsIndex
                                TCLResultRows = TCLResultRows + TclIndex
                                compareresxlssheet.write(calledRows - 2, 0, caserunxlssheet.cell_value(calledRows - 1, 1))
                                compareresxlssheet.write(calledRows - 2, 1, u"此条用例执行结果数与执行次数不符")
                            elif NmsIndex == TclIndex and NmsIndex <> int(RunNum):
                                print "two"
                                NmsresultRows = NmsresultRows + NmsIndex
                                TCLResultRows = TCLResultRows + TclIndex
                                compareresxlssheet.write(calledRows - 2, 0, caserunxlssheet.cell_value(calledRows - 1, 1))
                                compareresxlssheet.write(calledRows - 2, 1, u"此条用例执行结果数与执行次数不符")
                            elif NmsIndex == TclIndex and NmsIndex == int(RunNum):
                                print "same"
                                JudgeDataDetailTem1 = JudgeDataDetail
                                CallTypeData_backup = CallTypeData
                                CallModelData_backup = CallModelData
                                CallerNumData_backup = CallerNumData
                                CallerTypeData_backup = CallerTypeData
                                CalledNumData_backup = CalledNumData
                                CalledTypeData_backup = CalledTypeData
                                CallResultData_backup = CallResultData
                                while runnumIndex1 < int(RunNum):
                                    if TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1,TCLjudgedataColNum - 1) == "Pass":
                                        if len(JudgeDataDetailTem1) == 8:
                                            print WhetherServiceType
                                            print WhetherServiceTypeColNum
                                            if CallTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallTypeColNum).strip() and CallModelData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallModelColNum).strip() and CallerNumData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, CallerNumColNum).strip() and CallerTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallerTypeColNum).strip() and CalledNumData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CalledNumColNum).strip() and CalledTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CalledTypeColNum).strip() and CallResultData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, CallResultColNum).strip() and (NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, WhetherServiceTypeColNum).strip() == "是" or NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, WhetherServiceTypeColNum).strip() == "Yes"):
                                                compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1,  u"第" + str(runnumIndex1 + 1) + u"次执行结果一致")
                                                print NmsresultRows + runnumIndex1
                                                runnumIndex1 = runnumIndex1 + 1
                                                if runnumIndex1 < int(RunNum):
                                                    dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1,TCLjudgedataColNum).split('：')
                                                    servicetypeTem1 = dataarrTem1[0]
                                                    judgedataTem1 = dataarrTem1[1]
                                                    # print "下一条数据:"
                                                    # print judgedataTem1
                                                    JudgeDataDetailTem1 = judgedataTem1.split( '-')
                                                    CallTypeData_backup = JudgeDataDetailTem1[0]
                                                    CallModelData_backup = JudgeDataDetailTem1[1]
                                                    CallerNumData_backup = JudgeDataDetailTem1[2]
                                                    CallerTypeData_backup = JudgeDataDetailTem1[3]
                                                    CalledNumData_backup = JudgeDataDetailTem1[4]
                                                    CalledTypeData_backup = JudgeDataDetailTem1[5]
                                                    CallResultData_backup = JudgeDataDetailTem1[6]
                                                    if len(JudgeDataDetailTem1) == 8:
                                                        WhetherServiceType = JudgeDataDetailTem1[7]
                                                        RowNum, WhetherServiceTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, WhetherServiceType, 0, NmsResultxlscolumns)
                                            else:
                                                compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1, u"第" + str(runnumIndex1 + 1) + u"次执行结果不一致")
                                                print NmsresultRows + runnumIndex1
                                                runnumIndex1 = runnumIndex1 + 1
                                                if runnumIndex1 < int(RunNum):
                                                    dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1, TCLjudgedataColNum).split('：')
                                                    servicetypeTem1 = dataarrTem1[0]
                                                    judgedataTem1 = dataarrTem1[1]
                                                    JudgeDataDetailTem1 = judgedataTem1.split('-')

                                                    '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                                    for enlishi in range(0, len(JudgeDataDetailTem1)):
                                                        JudgeDataDetailTem1[enlishi] = JudgeDataDetailTem1[enlishi].replace("=", "-")
                                                    ''''''''''''''''''''''''''''''''''''''''''''

                                                    # print "下一条数据:"
                                                    # print judgedataTem1
                                                    CallTypeData_backup = JudgeDataDetailTem1[0]
                                                    CallModelData_backup = JudgeDataDetailTem1[1]
                                                    CallerNumData_backup = JudgeDataDetailTem1[2]
                                                    CallerTypeData_backup = JudgeDataDetailTem1[3]
                                                    CalledNumData_backup = JudgeDataDetailTem1[4]
                                                    CalledTypeData_backup = JudgeDataDetailTem1[5]
                                                    CallResultData_backup = JudgeDataDetailTem1[6]
                                                    if len(JudgeDataDetailTem1) == 8:
                                                        WhetherServiceType = JudgeDataDetailTem1[7]
                                                        RowNum, WhetherServiceTypeColNum = self.Func.GetCellRowAndCol(NmsResultxlssheet, WhetherServiceType, 0, NmsResultxlscolumns)
                                        elif len(JudgeDataDetailTem1) == 7:
                                            if CallTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallTypeColNum).strip() and CallModelData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallModelColNum).strip() and CallerNumData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CallerNumColNum).strip() and CallerTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, CallerTypeColNum).strip() and CalledNumData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CalledNumColNum).strip() and CalledTypeData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1,CalledTypeColNum).strip() and CallResultData_backup == NmsResultxlssheet.cell_value(NmsresultRows + runnumIndex1 - 1, CallResultColNum).strip():
                                                compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1, u"第" + str(runnumIndex1 + 1) + u"次执行结果一致")
                                                print NmsresultRows + runnumIndex1
                                                runnumIndex1 = runnumIndex1 + 1
                                                if runnumIndex1 < int(RunNum):
                                                    dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1,TCLjudgedataColNum).split('：')
                                                    servicetypeTem1 = dataarrTem1[0]
                                                    judgedataTem1 = dataarrTem1[1]
                                                    JudgeDataDetailTem1 = judgedataTem1.split('-')

                                                    '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                                    for enlishi in range(0, len(JudgeDataDetailTem1)):
                                                        JudgeDataDetailTem1[enlishi] = JudgeDataDetailTem1[enlishi].replace("=", "-")
                                                    ''''''''''''''''''''''''''''''''''''''''''''


                                                    # print "下一条数据:"
                                                    # print judgedataTem1
                                                    CallTypeData_backup = JudgeDataDetailTem1[0]
                                                    CallModelData_backup = JudgeDataDetailTem1[1]
                                                    CallerNumData_backup = JudgeDataDetailTem1[2]
                                                    CallerTypeData_backup = JudgeDataDetailTem1[3]
                                                    CalledNumData_backup = JudgeDataDetailTem1[4]
                                                    CalledTypeData_backup = JudgeDataDetailTem1[5]
                                                    CallResultData_backup = JudgeDataDetailTem1[6]
                                            else:
                                                compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1, u"第" + str(runnumIndex1 + 1) + u"次执行结果不一致")
                                                print NmsresultRows + runnumIndex1
                                                runnumIndex1 = runnumIndex1 + 1
                                                if runnumIndex1 < int(RunNum):
                                                    dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1,TCLjudgedataColNum).split('：')
                                                    servicetypeTem1 = dataarrTem1[0]
                                                    judgedataTem1 = dataarrTem1[1]
                                                    JudgeDataDetailTem1 = judgedataTem1.split('-')

                                                    '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                                    for enlishi in range(0, len(JudgeDataDetailTem1)):
                                                        JudgeDataDetailTem1[enlishi] = JudgeDataDetailTem1[enlishi].replace("=", "-")
                                                    ''''''''''''''''''''''''''''''''''''''''''''


                                                    # print "下一条数据:"
                                                    # print judgedataTem1
                                                    CallTypeData_backup = JudgeDataDetailTem1[0]
                                                    CallModelData_backup = JudgeDataDetailTem1[1]
                                                    CallerNumData_backup = JudgeDataDetailTem1[2]
                                                    CallerTypeData_backup = JudgeDataDetailTem1[3]
                                                    CalledNumData_backup = JudgeDataDetailTem1[4]
                                                    CalledTypeData_backup = JudgeDataDetailTem1[5]
                                                    CallResultData_backup = JudgeDataDetailTem1[6]
                                    else:
                                        compareresxlssheet.write(calledRows - 2, runnumIndex1 + 1, u"TCL第" + str(runnumIndex1 + 1) + u"次执行结果失败，不进行比较")
                                        print NmsresultRows + runnumIndex1
                                        runnumIndex1 = runnumIndex1 + 1
                                NmsresultRows = NmsresultRows + int(RunNum)
                                TCLResultRows = TCLResultRows + int(RunNum)
                        else:
                            TCLResultRows = TCLResultRows + 1
                    elif servicetype == "数据":
                        if D_NmsresultRows < D_NmsResultxlsrows + 1:
                            JudgeDataDetail = judgedata.split('-')

                            '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                            for enlishi in range(0, len(JudgeDataDetail)):
                                JudgeDataDetail[enlishi] = JudgeDataDetail[enlishi].replace("=", "-")
                            ''''''''''''''''''''''''''''''''''''''''''''

                            print judgedata
                            D_CallTypeData = JudgeDataDetail[0]
                            D_CallerNumData = JudgeDataDetail[1]
                            D_CallerTypeData = JudgeDataDetail[2]
                            D_CalledNumData = JudgeDataDetail[3]
                            D_CalledTypeData = JudgeDataDetail[4]
                            D_SendResultData = JudgeDataDetail[5]
                            calltypesplit = caserunxlssheet.cell_value(calledRows - 1, 1).split('（')
                            print caserunxlssheet.cell_value(calledRows - 1, 1)
                            caseruncalltype = calltypesplit[0]
                            print caseruncalltype
                            print D_CallTypeData
                            D_CallTypeData_backup = D_CallTypeData
                            D_NmsIndex = 0
                            D_TclIndex = 0
                            while D_NmsresultRows < D_NmsResultxlsrows + 1:
                                if D_NmsresultRows + D_NmsIndex - 1 < D_NmsResultxlsrows:
                                    if caseruncalltype == D_NmsResultxlssheet.cell_value(D_NmsresultRows + D_NmsIndex - 1, D_CallTypeColNum).strip():
                                        D_NmsIndex = D_NmsIndex + 1
                                    else:
                                        break
                                else:
                                    break
                            while TCLResultRows < TCLResultxlsrows + 1:
                                if caseruncalltype == D_CallTypeData_backup:
                                    D_TclIndex = D_TclIndex + 1
                                    if TCLResultRows + D_TclIndex < TCLResultxlsrows + 1:
                                        dataarrTem = TCLResultxlssheet.cell_value(TCLResultRows + D_TclIndex - 1, TCLjudgedataColNum).split('：')
                                        servicetypeTem = dataarrTem[0]
                                        judgedataTem = dataarrTem[1]
                                        # print "下一条数据:"
                                        # print judgedataTem
                                        JudgeDataDetailTem = judgedataTem.split('-')

                                        '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                        for enlishi in range(0, len(JudgeDataDetailTem)):
                                            JudgeDataDetailTem[enlishi] = JudgeDataDetailTem[enlishi].replace("=", "-")
                                        ''''''''''''''''''''''''''''''''''''''''''''

                                        D_CallTypeData_backup = JudgeDataDetailTem[0]
                                    else:
                                        break
                                else:
                                    break
                            print D_NmsIndex
                            print D_TclIndex
                            if D_NmsIndex <> D_TclIndex:
                                print "one"
                                D_NmsresultRows = D_NmsresultRows + D_NmsIndex
                                TCLResultRows = TCLResultRows + D_TclIndex
                                compareresxlssheet.write(calledRows - 2, 0, caserunxlssheet.cell_value(calledRows - 1, 1))
                                compareresxlssheet.write(calledRows - 2, 1, u"此条用例执行结果数与执行次数不符")
                            elif D_NmsIndex == D_TclIndex and D_NmsIndex <> int(D_RunNum):
                                print "two"
                                D_NmsresultRows = D_NmsresultRows + D_NmsIndex
                                TCLResultRows = TCLResultRows + D_TclIndex
                                compareresxlssheet.write(calledRows - 2, 0, caserunxlssheet.cell_value(calledRows - 1, 1))
                                compareresxlssheet.write(calledRows - 2, 1, u"此条用例执行结果数与执行次数不符")
                            elif D_NmsIndex == D_TclIndex and D_NmsIndex == int(D_RunNum):
                                print "same"
                                JudgeDataDetailTem1 = JudgeDataDetail
                                D_CallTypeData_backup = D_CallTypeData
                                D_CallerNumData_backup = D_CallerNumData
                                D_CallerTypeData_backup = D_CallerTypeData
                                D_CalledNumData_backup = D_CalledNumData
                                D_CalledTypeData_backup = D_CalledTypeData
                                D_SendResultData_backup = D_SendResultData
                                while runnumIndex1 < int(D_RunNum):
                                    if TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1,TCLjudgedataColNum - 1) == "Pass":
                                        if D_CallTypeData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1,D_CallTypeColNum).strip() and D_CallerNumData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1,D_CallerNumColNum).strip() and D_CallerTypeData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1,D_CallerTypeColNum).strip() and D_CalledNumData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1, D_CalledNumColNum).strip() and D_CalledTypeData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1,D_CalledTypeColNum).strip() and D_SendResultData_backup == D_NmsResultxlssheet.cell_value(D_NmsresultRows + runnumIndex1 - 1,D_SendResultColNum).strip():
                                            compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1, u"第" + str(runnumIndex1 + 1) + u"次执行结果一致")
                                            #print D_NmsresultRows + runnumIndex1
                                            runnumIndex1 = runnumIndex1 + 1
                                            if runnumIndex1 < int(D_RunNum):
                                                dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1, TCLjudgedataColNum).split('：')
                                                servicetypeTem1 = dataarrTem1[0]
                                                judgedataTem1 = dataarrTem1[1]
                                                JudgeDataDetailTem1 = judgedataTem1.split('-')

                                                '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                                for enlishi in range(0, len(JudgeDataDetailTem1)):
                                                    JudgeDataDetailTem1[enlishi] = JudgeDataDetailTem1[enlishi].replace("=", "-")
                                                ''''''''''''''''''''''''''''''''''''''''''''

                                                # print "下一条数据:"
                                                # print judgedataTem1
                                                D_CallTypeData_backup = JudgeDataDetailTem1[0]
                                                D_CallerNumData_backup = JudgeDataDetailTem1[1]
                                                D_CallerTypeData_backup = JudgeDataDetailTem1[2]
                                                D_CalledNumData_backup = JudgeDataDetailTem1[3]
                                                D_CalledTypeData_backup = JudgeDataDetailTem1[4]
                                                D_SendResultData_backup = JudgeDataDetailTem1[5]
                                        else:
                                            compareresxlssheet.write(calledRows - 2,runnumIndex1 + 1, u"第" + str(runnumIndex1) + 1 + u"次执行结果不一致")
                                            #print D_NmsresultRows + runnumIndex1
                                            runnumIndex1 = runnumIndex1 + 1
                                            if runnumIndex1 < int(D_RunNum):
                                                dataarrTem1 = TCLResultxlssheet.cell_value(TCLResultRows + runnumIndex1 - 1, TCLjudgedataColNum).split('：')
                                                servicetypeTem1 = dataarrTem1[0]
                                                judgedataTem1 = dataarrTem1[1]
                                                JudgeDataDetailTem1 = judgedataTem1.split('-')

                                                '''英文版网管 Inter=BS Call --> Inter-BS Call'''
                                                for enlishi in range(0, len(JudgeDataDetailTem1)):
                                                    JudgeDataDetailTem1[enlishi] = JudgeDataDetailTem1[enlishi].replace("=", "-")
                                                ''''''''''''''''''''''''''''''''''''''''''''

                                                # print "下一条数据:"
                                                # print judgedataTem1
                                                D_CallTypeData_backup = JudgeDataDetailTem1[0]
                                                D_CallerNumData_backup = JudgeDataDetailTem1[1]
                                                D_CallerTypeData_backup = JudgeDataDetailTem1[2]
                                                D_CalledNumData_backup = JudgeDataDetailTem1[3]
                                                D_CalledTypeData_backup = JudgeDataDetailTem1[4]
                                                D_SendResultData_backup = JudgeDataDetailTem1[5]
                                    else:
                                        compareresxlssheet.write(calledRows - 2, runnumIndex1 + 1, u"TCL第" + str(runnumIndex1 + 1) + u"次执行结果失败，不进行比较")
                                        print D_NmsresultRows + runnumIndex1
                                        runnumIndex1 = runnumIndex1 + 1
                                D_NmsresultRows = D_NmsresultRows + int(D_RunNum)
                                TCLResultRows = TCLResultRows + int(D_RunNum)
                        else:
                            break
                else:
                    print "TCL_LOG end"
                    break
            calledRows = calledRows + 1
        self.Func.del_files(D_fileaddr, CompareResultXls)
        compareresxlsbook.save(str(D_fileaddr) + '\\' + str(CompareResultXls))
        driver.quit()


    def System_Statistic(self):
        call_count = {}
        call_type = ['Group Call,','Individual Call,','Inter-BS,','Intra-BS,','Intra-system,','Inter-system,','Normal Call','GPS Group Call','Include Call','Patch Group Call','PSTN Call','Reserved Call','Broadcast Call','Full-duplex Call','Emergency Call','Encrypted Call','Ambience Listening']
        data_type = ['Group Call','Individual Call','Individual Short Message Call','Individual Packet Data Call','Individual Status Message Call','Group Short Message Call','Group Packet Data Call','Group Status Message Call']
        # self.Func.csv_to_xlsx_pd('E:\NM_web_selenium_all\StatisticsCompare\TCL_LOG.csv','E:\NM_web_selenium_all\StatisticsCompare\TCL_LOG.xls')
        # self.Func.log2format('E:\NM_web_selenium_all\StatisticsCompare\TCL_LOG.xls','E:\NM_web_selenium_all\StatisticsCompare\Format_LOG.xls')
        excel_path = 'E:\NM_web_selenium_all\StatisticsCompare\Format_LOG.xls'
        d = pd.read_excel(excel_path, sheet_name='format_check_data', encoding='gbk',keep_default_na=False)
        #
        #
        # outfile = d[(d[u'Result']=='Pass')&(d[u'SpecialType'] == '加密呼叫')]#筛选
        data_LogData={}
        data_count = {}
        #数据组呼次数：Result为'pass'，ServiceType为'数据'，CallType为'短消息组呼'或'状态消息组呼'或'长短消息组呼'
        data_LogData[data_type[0]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&((d[u'CallType']=='短消息组呼')|(d[u'CallType']=='状态消息组呼')|(d[u'CallType']=='长短消息组呼'))]
        #数据单呼次数：Result为'pass'，ServiceType为'数据'，CallType为'短消息单呼'或'状态消息单呼'或'长短消息单呼'
        data_LogData[data_type[1]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&((d[u'CallType']=='短消息单呼')|(d[u'CallType']=='状态消息单呼')|(d[u'CallType']=='长短消息单呼'))]
        #短信单呼次数：Result为'pass'，ServiceType为'数据'，CallType为'短消息单呼'
        data_LogData[data_type[2]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='短消息单呼')]
        #长短信单呼次数：Result为'pass'，ServiceType为'数据'，CallType为'长短消息单呼'
        data_LogData[data_type[3]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='长短消息单呼')]
        #状态消息单呼次数：Result为'pass'，ServiceType为'数据'，CallType为'长短消息单呼'
        data_LogData[data_type[4]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='状态消息单呼')]
        #短信组呼次数：Result为'pass'，ServiceType为'数据'，CallType为'短消息单呼'
        data_LogData[data_type[5]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='短消息组呼')]
        #长短信组呼次数：Result为'pass'，ServiceType为'数据'，CallType为'长短消息单呼'
        data_LogData[data_type[6]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='长短消息组呼')]
        #状态消息组呼次数：Result为'pass'，ServiceType为'数据'，CallType为'长短消息单呼'
        data_LogData[data_type[7]] =d[(d[u'Result']=='Pass')&(d[u'ServiceType']=='数据')&(d[u'CallType']=='状态消息组呼')]
        #各种data——type呼叫次数
        for i in range(len(data_LogData)):
            data_count[data_type[i]] = len(data_LogData[data_type[i]])
        #数据单呼呼叫次数:
        di = data_LogData[data_type[1]]  # 传入筛选出来的单呼数据，返回每个msi的次数
        end_result_in = self.Func.get_indi_count(di,type = 'msi')
        #数据组呼呼叫次数
        di = data_LogData[data_type[0]]  # 传入筛选出来的单呼数据，返回每个msi的次数
        end_result_gr = self.Func.get_indi_count(di, type='gsi')
        #
        #     print di[(di[u'CallingAddr']==u'3131001')]
            # | (di[u'CalledAddr'] == unicode(str(msi_num_calling)))


        # print data_LogData[data_type[1]][u'CallingAddr'].drop_duplicates()

