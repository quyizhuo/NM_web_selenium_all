# -*- coding: utf-8 -*-

"""
@version: 
@author: lidan 
@file: main.py
@time: 2018/1/16 17:02
"""
from EOES import EOES
from SmallFunctionsOfNM import SmallFuncionsOfNM
from func import func
class main:
    if __name__=="__main__":
        func = func()
        eoes = EOES()
        smallfunctionsofnm = SmallFuncionsOfNM()
        eoes.System_Statistic()
        # eoes.DetailedListStatistics()
        # eoes.Customizable_Statistics()
        # eoes.System_BS_Organization_Statistics()
        # eoes.KPI_Statistic()





        # smallfunctionsofnm.CrossSystemConfig()
        # smallfunctionsofnm.updateChannelList()
        # smallfunctionsofnm.add_VPN()
        # smallfunctionsofnm.delete_VPN()
        # smallfunctionsofnm.FaultWeakeningPlan()
        # func.CloseProcess('EXCEL,chromedriver,conhost') # 最后执行