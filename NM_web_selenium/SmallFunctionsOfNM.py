# -*- coding: utf-8 -*-

"""
@version: 
@author: lidan 
@file: SmallFunctionsOfNM.py
@time: 2018/2/27 11:49
"""

from selenium import webdriver
from func import func
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from itertools import combinations
import sys
reload(sys)
sys.setdefaultencoding('utf8')
class SmallFuncionsOfNM:
    def __init__(self):
        self.Func = func()

    # 循环添加VPN节点
    def add_VPN(self):
        # 从配置文件获取登录参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))  # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        # 从配置文件获取VPN参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet2')
        OneNodeNum = str(configsheet.cell_value(2, 1))  # 添加一级节点数
        OneNodeOfAddTowNode = str(configsheet.cell_value(2, 2))  # 添加二级节点的一级节点数
        TowNodeNum = str(configsheet.cell_value(2, 3))  # 每个一级节点添加二级节点数
        driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]').click() # 点击用户
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/span').click() # 点击用户管理
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/div/div[1]').click() # 点击本地用户
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[5]/div[1]/div[2]/a', 3).click() # 点击展开VPN箭头

        for onenode in range(0, int(OneNodeNum), 1):
            self.Func.find_element_by_xpath(driver, '//*[@id="zTreeVPN_1_span"]',1).click()  # 点击“组织架构”一级节点
            self.Func.find_element_by_xpath(driver, '//*[@id="addSonNode"]').click()  # 添加子节点按钮
            OneNodeName = 'zdh_' + str(onenode + 1)
            print "add " + OneNodeName
            self.Func.find_element_by_xpath(driver, '//*[@id="addNodeName"]').send_keys(OneNodeName)  # 子节点名称
            self.Func.find_element_by_xpath(driver, '//*[@id="addNodeCode"]').send_keys(OneNodeName)  # 子节点代码
            self.Func.find_element_by_xpath(driver, '//*[@id="addNodeCommit"]').click() # 点击提交
            self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
            if onenode < int(OneNodeOfAddTowNode):
                for townode in range(0, int(TowNodeNum), 1):
                    self.Func.find_element_by_css_selector(driver, 'a[title='+ OneNodeName + ']', 1).click()
                    self.Func.find_element_by_xpath(driver, '//*[@id="addSonNode"]').click()  # 添加子节点按钮
                    TwoNodeName = OneNodeName + '_' + str(townode + 1)
                    print "    add " + TwoNodeName
                    self.Func.find_element_by_xpath(driver, '//*[@id="addNodeName"]').send_keys(TwoNodeName)  # 子节点名称
                    self.Func.find_element_by_xpath(driver, '//*[@id="addNodeCode"]').send_keys(TwoNodeName)  # 子节点代码
                    self.Func.find_element_by_xpath(driver, '//*[@id="addNodeCommit"]').click()  # 点击提交
                    self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
                time.sleep(2)
                driver.quit()
                driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
                self.Func.login(driver, username, password)  # 登录网管
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]').click()  # 点击用户
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/span').click()  # 点击用户管理
                self.Func.find_element_by_xpath(driver,'//*[@id="left_menu"]/div[4]/div/div[1]/div/div[1]').click()  # 点击本地用户
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[5]/div[1]/div[2]/a',3).click()  # 点击展开VPN箭头
        time.sleep(2)
        driver.quit()

    # 删除添加的VPN节点
    def delete_VPN(self):
        # 从配置文件获取登录参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0)) # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        # 从配置文件获取VPN参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet2')
        OneNodeNum = str(configsheet.cell_value(2, 1))  # 添加一级节点数
        OneNodeOfAddTowNode = str(configsheet.cell_value(2, 2))  # 添加二级节点的一级节点数
        TowNodeNum = str(configsheet.cell_value(2, 3))  # 每个一级节点添加二级节点数
        driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]').click() # 点击用户
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/span').click() # 点击用户管理
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/div/div[1]').click() # 点击本地用户
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[5]/div[1]/div[2]/a',5).click() # 点击展开VPN箭头
        for onenode in range(0, int(OneNodeNum), 1):
            OneNodeName = 'zdh_' + str(onenode + 1)
            if onenode < int(OneNodeOfAddTowNode):
                for townode in range(0, int(TowNodeNum), 1):
                    TwoNodeName = OneNodeName + '_' + str(townode + 1)
                    try:
                        print "delete " + TwoNodeName
                        self.Func.find_element_by_css_selector(driver, 'a[title='+ TwoNodeName + ']', 1).click() # 点击二级节点
                        self.Func.find_element_by_xpath(driver, '//*[@id="treeDelete"]').click()  # 删除节点按钮
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0', 0.5).click()  # 点击确认
                        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0', 0.5).click()  # 点击确定
                    except:
                        print "  delete " + TwoNodeName + " wrong"
                        pass
                time.sleep(2)
                driver.quit()
                driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
                self.Func.login(driver, username, password)  # 登录网管
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]').click()  # 点击用户
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[4]/div/div[1]/span').click()  # 点击用户管理
                self.Func.find_element_by_xpath(driver,'//*[@id="left_menu"]/div[4]/div/div[1]/div/div[1]').click()  # 点击本地用户
                self.Func.switch_to_frame(driver, 'mainFrame')
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[5]/div[1]/div[2]/a',3).click()  # 点击展开VPN箭头
            try:
                print "    delete " + OneNodeName
                self.Func.find_element_by_css_selector(driver, 'a[title=' + OneNodeName + ']', 1).click()  # 点击一级节点
                self.Func.find_element_by_xpath(driver, '//*[@id="treeDelete"]').click()  # 删除节点按钮
                self.Func.find_element_by_class_name(driver, 'layui-layer-btn0', 0.5).click()  # 点击确认
                self.Func.find_element_by_class_name(driver, 'layui-layer-btn0', 0.5).click()  # 点击确定
            except:
                print "      delete " + OneNodeName + " wrong"
                pass
        driver.quit()

    # 更新信道列表
    def updateChannelList(self):
        # 从配置文件获取登录参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))  # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        # 从配置文件获取"更新信道列表"参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet2')
        TemplateNum = int(configsheet.cell_value(6, 1))  # 模板编号及名称
        TSCNumSE = str(configsheet.cell_value(6, 2))  # 基站起止编号
        ChannelMachineNumSE = str(configsheet.cell_value(6, 3))  # 信道机起止编号
        DatumFrequency = str(configsheet.cell_value(6, 4))# 发射基准频点
        LaunchFrequency = str(configsheet.cell_value(6, 5))  # 信道机发射频点
        StartLAI = int(configsheet.cell_value(6, 6))  # LAI开始编号
        FrequencyIncrement = 0.000001 # 频点增量
        StartTSCNum, EndTSCNum = map(int,TSCNumSE.split('-'))
        StartChannelMachineNum, EndChannelMachineNum = map(int,ChannelMachineNumSE.split('-'))

        driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]').click()  # 点击配置
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]/div/div[7]').click()  # 点击更新信道列表
        self.Func.switch_to_frame(driver, 'mainFrame')
        frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[1]/div/iframe', 1)
        self.Func.switch_to_frame(driver, 'mainFrame')
        driver.switch_to_frame(frame)
        self.Func.find_element_by_xpath(driver, '//*[@id="mbcj"]/i').click() # 点击添加模板
        self.Func.switch_to_frame(driver, 'mainFrame')
        driver.switch_to_frame(frame)
        driver.switch_to_frame('layui-layer-iframe1')
        self.Func.find_element_by_xpath(driver, '//*[@id="templateNo"]').send_keys(TemplateNum)  # 模板编号
        self.Func.find_element_by_xpath(driver, '//*[@id="templateName"]').send_keys(TemplateNum)  # 模板名称
        self.Func.find_element_by_xpath(driver, '//*[@id="channelTabs"]/div[1]/div[3]/ul/li[2]/a/span[1]').click()  # 点击未知信道信息
        j = 0
        frameno = 1
        for TSC in range(StartTSCNum, EndTSCNum + 1, 1):
            i = 0
            print "add TSC " + str(TSC)
            for ChannelMachineNum in range(StartChannelMachineNum, EndChannelMachineNum + 1, 1):
                self.Func.find_element_by_xpath(driver, '//*[@id="tjxg"]/i', 1).click() # 点击添加基站编号按钮
                self.Func.switch_to_frame(driver, 'mainFrame')
                driver.switch_to_frame(frame)
                driver.switch_to_frame('layui-layer-iframe1')
                time.sleep(0.5)
                driver.switch_to_frame('layui-layer-iframe' + str(frameno))
                frameno = frameno + 1
                self.Func.find_element_by_xpath(driver, '//*[@id="station"]').send_keys(TSC)  # 输入基站编号
                self.Func.find_element_by_xpath(driver, '//*[@id="save"]').click()  # 点击确认
                self.Func.switch_to_frame(driver, 'mainFrame')
                driver.switch_to_frame(frame)
                driver.switch_to_frame('layui-layer-iframe1')
                self.Func.find_element_by_xpath(driver, '//*[@id="tscNameStr"]').send_keys(TSC)  # 基站名称
                self.Func.find_element_by_xpath(driver, '//*[@id="channelID"]').send_keys(ChannelMachineNum)  # 信道机编号
                self.Func.find_element_by_xpath(driver, '//*[@id="downFreqFloat"]').send_keys(DatumFrequency)  # 发射基准频点
                self.Func.find_element_by_xpath(driver, '//*[@id="txFreqFloat"]').send_keys(str(float(LaunchFrequency) + FrequencyIncrement * i))  # 信道机发射频点
                self.Func.find_element_by_xpath(driver, '//*[@id="lai"]').send_keys(StartLAI + j)  # LAI
                i = i + 1
                self.Func.find_element_by_xpath(driver, '//*[@id="kh"]/div[2]/div/div[1]/div/div[1]/div[2]/i').click() # 点击添加信道信息
            j = j + 1

        self.Func.find_element_by_xpath(driver, '//*[@id="save"]', 1).click() # 点击确认
        time.sleep(2)
        driver.quit()

    # 跨系统配置
    def CrossSystemConfig(self):
        # 从配置文件获取登录参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))  # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        # 从配置文件获取"更新信道列表"参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet2')
        SystemId = int(configsheet.cell_value(10, 1))  # 系统号
        ThisSystemAreaCode = str(configsheet.cell_value(10, 2))  # 本系统区号
        TeamNoES = str(configsheet.cell_value(10, 3))  # 起止队号

        StartTeamNo, EndTeamNo = map(int, TeamNoES.split('-'))
        ThisSystemAreaCodeList = map(int, ThisSystemAreaCode.split(','))

        driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]').click()  # 点击配置
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]/div/div[2]').click()  # 点击跨系统配置
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div[1]/div/div/div[1]/div[1]/i').click() # 添加系统号
        self.Func.switch_to_frame(driver, 'mainFrame,layui-layer-iframe1')
        self.Func.find_element_by_xpath(driver, '//*[@id="sysNo"]').send_keys(SystemId) # 输入系统号
        self.Func.find_element_by_xpath(driver, '/html/body/div[2]/div/div/a[1]').click() # 点击保存
        self.Func.switch_to_frame(driver, 'mainFrame')
        self.Func.find_element_by_class_name(driver, 'layui-layer-btn0').click()  # 点击确定
        areaNum = 0
        i = 1
        for areano in range(0, 100, 1):
            if areano in ThisSystemAreaCodeList:
                continue
            else:
                if areano < 10:
                    areano1 = '0' + str(areano)
                else:
                    areano1 = str(areano)
                print "add areano " + str(areano1)
                for teamno in range(StartTeamNo, EndTeamNo + 1, 1):
                    if teamno < 10:
                        teamno1 = '0' + str(teamno)
                    else:
                        teamno1 = str(teamno)
                    print "    add teamno " + teamno1
                    self.Func.switch_to_frame(driver, 'mainFrame')
                    self.Func.find_element_by_xpath(driver, '//input[@value=' + str(SystemId) + ']', 1).click()  # 选择刚刚添加的系统号
                    self.Func.switch_to_frame(driver, 'mainFrame,acrossFrame')
                    self.Func.find_element_by_xpath(driver, '/html/body/div/div[1]/div[1]/i').click()  # 点击添加段对信息
                    self.Func.switch_to_frame(driver, 'mainFrame,layui-layer-iframe' + str(i))
                    i = i + 1
                    self.Func.find_element_by_xpath(driver, '//*[@id="addrScopeForm"]/table/tbody/tr[1]/td[2]/input').clear()
                    self.Func.find_element_by_xpath(driver, '//*[@id="addrScopeForm"]/table/tbody/tr[1]/td[2]/input').send_keys(str(areano1))  # 输入区号
                    self.Func.find_element_by_xpath(driver, '//*[@id="gh"]').click()  # 勾选“个呼队号配置”
                    self.Func.Select(driver, '//*[@id="addrScopeForm"]/table/tbody/tr[3]/td[1]/span/span/span','//*[@id="_easyui_combobox_i1_' + str(teamno) + '"]')  # 选择起始队号
                    self.Func.Select(driver, '//*[@id="addrScopeForm"]/table/tbody/tr[4]/td[1]/span/span/span','//*[@id="_easyui_combobox_i3_' + str(teamno) + '"]')  # 选择截止队号
                    self.Func.find_element_by_xpath(driver, '/html/body/div[2]/div/div/a[1]').click() # 点击保存
                    self.Func.find_element_by_xpath(driver, '//*[@id="layui-layer1"]/div[3]/a', 1).click() # 点击确定
            areaNum = areaNum + 1
            if areaNum % 10 == 0: # 设置每添加10个区号重新登录浏览器
                i = 1
                driver.quit()
                driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
                self.Func.login(driver, username, password)  # 登录网管
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]', 1).click()  # 点击配置
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]/div/div[2]').click()  # 点击跨系统配置
        time.sleep(2)
        driver.quit()

    # 故障弱化预案
    def FaultWeakeningPlan(self):
        # 从配置文件获取登录参数'//*[@id="datagrid-row-r1-2-0"]/td[2]/div'
        configsheet = self.Func.read_excel('config.xls', 'Sheet1')
        loginPageAddr = str(configsheet.cell_value(1, 0))  # 登录地址
        username = str(configsheet.cell_value(1, 1))  # 用户名
        password = str(configsheet.cell_value(1, 2))  # 密码
        # 从配置文件获取"更新信道列表"参数
        configsheet = self.Func.read_excel('config.xls', 'Sheet2')
        PlanName = configsheet.cell_value(14, 1)  # 预案组名称
        print u"预案组名称：" + str(PlanName)
        TSBSClist = []  # 可选群组基站（32个）
        for row in range(14, 46, 1):
            TSBSClist.append(str(configsheet.cell_value(row, 2)))
        print  u"可选基站个数：" + str(len(TSBSClist)), TSBSClist
        driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
        self.Func.login(driver, username, password)  # 登录网管
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]').click()  # 点击配置
        self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]/div/div[6]').click()  # 故障弱化预案
        self.Func.switch_to_frame(driver, 'mainFrame')
        frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[1]/div/iframe', 1)
        self.Func.switch_to_frame(driver, 'mainFrame')
        driver.switch_to_frame(frame)
        self.Func.find_element_by_xpath(driver, '//input[@value=' + str(PlanName) + ']', 1).click()  # 选择预案组
        self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div/div/div/div[5]/i').click()  # 点击查看预案组预案
        PlanNoList = list(combinations(TSBSClist, 1))
        PlanNoList = PlanNoList + list(combinations(TSBSClist, 2))
        PlanNoList = PlanNoList + list(combinations(TSBSClist, 3))
        PlanNo = 0
        while PlanNo < 1024:
            print PlanNo, PlanNoList[PlanNo]
            try:
                self.Func.switch_to_frame(driver, 'mainFrame')
                frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[2]/div/iframe', 1)
                self.Func.switch_to_frame(driver, 'mainFrame')
                driver.switch_to_frame(frame)
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div/div/div/div[2]/i').click() # 点击新增预案组预案
                iframes = driver.find_elements_by_tag_name('iframe')
                for iframe in iframes:
                    if 'layui-layer-iframe' in iframe.get_property('id'):
                        self.Func.switch_to_frame(driver, 'mainFrame')
                        frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[2]/div/iframe', 1)
                        self.Func.switch_to_frame(driver, 'mainFrame')
                        driver.switch_to_frame(frame)
                        driver.switch_to_frame(iframe.get_property('id'))
                for TSBSC in PlanNoList[PlanNo]:
                    self.Func.find_element_by_xpath(driver, '//div[contains(text(),"' + TSBSC + '")]').click()  # 选择基站
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div/table[1]/tbody/tr/td[2]/a').click()# 点击添加基站
                self.Func.find_element_by_xpath(driver, '/html/body/div[2]/div/div[1]/a[1]').click() # 点击下一步
                self.Func.find_element_by_xpath(driver, '//div[contains(text(),"群组基站")]').click()  # 任选一行
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div/table[2]/tbody/tr/td[2]/div/div[2]/div[1]/table/tbody/tr/td/a/span').click()  #　点击修改按钮
                self.Func.find_element_by_xpath(driver, '//*[@id="txPw"]').send_keys('1') # 输入发射功率
                self.Func.find_element_by_xpath(driver, '//*[@id="bs"]/div/a[1]').click()  # 点击保存
                self.Func.find_element_by_xpath(driver, '/html/body/div[2]/div/div[2]/a[2]').click()  # 点击保存
                self.Func.switch_to_frame(driver, 'mainFrame')
                frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[2]/div/iframe', 0.5)
                self.Func.switch_to_frame(driver, 'mainFrame')
                driver.switch_to_frame(frame)
                self.Func.find_element_by_class_name(driver, 'layui-layer-btn0', 0.5).click()  # 点击确定
                PlanNo = PlanNo + 1
            except Exception, e:
                print u"添加第" + str(PlanNo + 1) + u"个预案错误，将重新添加!"
                print str(e)
                driver.quit()
                driver = self.Func.OpenChrome(loginPageAddr)  # 打开chrome
                self.Func.login(driver, username, password)  # 登录网管
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]').click()  # 点击配置
                self.Func.find_element_by_xpath(driver, '//*[@id="left_menu"]/div[2]/div/div[6]').click()  # 故障弱化预案
                self.Func.switch_to_frame(driver, 'mainFrame')
                frame = self.Func.find_element_by_xpath(driver, '//*[@id="tt"]/div[2]/div[1]/div/iframe', 1)
                self.Func.switch_to_frame(driver, 'mainFrame')
                driver.switch_to_frame(frame)
                self.Func.find_element_by_xpath(driver, '//input[@value=' + str(PlanName) + ']', 1).click()  # 选择预案组
                self.Func.find_element_by_xpath(driver, '/html/body/div[1]/div/div/div/div[5]/i').click()  # 点击查看预案组预案
        driver.quit()







