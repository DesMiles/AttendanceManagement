import os
import wx
import wx.grid
import webbrowser
import datetime
import shutil
import random
import time
import xlsxwriter
import json
import csv
import cv2
import FaceRecognize
import ApplicationTool
from skimage import io as iio
from importlib import reload

#功能菜单按钮ID
ID_Service_Start = 1
ID_Service_Close = 2
ID_ShowEmployee = 3
ID_ShowTimeInfo = 4
ID_SettingTime = 5
ID_SettingLeaveTime = 6

#考勤表菜单按钮ID
ID_SearchLogList = 11
ID_SearchLogList_Test = 12
ID_CloseLogList = 13
ID_ExportLogList = 14
ID_UpLoad_LogList = 15
ID_DownLoad_LogList = 16

#测试菜单按钮ID
ID_Register = 21
ID_PunchCard = 22
ID_Leave = 23
ID_AddFaceData = 24
ID_AddPunchCardInfo_Month = 25
ID_AddPunchCardInfo_Year = 26

#帮助菜单按钮ID
ID_InitData = 31
ID_Help = 32
ID_About = 33

#主背景图路径
MainPic = "Picture/bmp.png"
#数据库路径
path_database = "./Data/loglist.db"
#人脸样本保存路径
path_face_sample = "./Data/face_img_database/"
#人脸库数据保存路径
path_feature_all = "./Data/feature_all.csv"
#存放设置打卡时间的数据路径
path_setting_time = "./Data/setting_time.json"
#人脸识别分类器路径
path_classifiers = "Classifiers/haarcascade_frontalface_alt2.xml"
#云服务器IP地址及端口号
IP = '8.129.64.78'
Port = 22
#云服务器存储地址
path_remote = "/home/CloudDataBase/"

class MainUI(wx.Frame):
    def __init__(self,rut):
        wx.Frame.__init__(self,parent=rut,title="打卡系统总控制台(V2.8)",size=(1200,660),style=wx.CAPTION|wx.CLOSE_BOX)
        wx.Frame.SetMinSize(self,size=(1200,660))
        wx.Frame.SetMaxSize(self,size=(1220,660))
        self.SetBackgroundColour('white')

        self.database = ApplicationTool.DataBase(path_database)
        self.CreateInfoText()
        self.CreateGallery()
        self.aispeech = ApplicationTool.AISpeech()

        #人脸录入计时器
        self.count = 0
        #打卡人员工号初始化
        self.ID = ''
        #人脸录入姓名标签初始化
        self.name = ''
        self.name_info = ''
        #人脸录入处理重复录入标志（False代表无重复）
        self.repeat_register_flag = False
        #获取本地存储的打卡设置时间（如果无设置则默认为07:00:00~09:00:00为上班打卡时间，17:00:00为下班打卡时间）
        with open(path_setting_time, 'r', encoding='utf8') as fp:
            TimeData = json.load(fp)
            self.MinTime = TimeData.get('MinTime')
            self.MaxTime = TimeData.get('MaxTime')
            self.LeavingTime = TimeData.get('LeavingTime')

        #生成菜单栏
        MenuBar = wx.MenuBar()
        menu_Font = wx.Font()
        menu_Font.SetPointSize(12)
        menu_Font.SetWeight(wx.BOLD)

        #生成功能菜单
        FunctionMenu = wx.Menu()
        self.Service_Start = wx.MenuItem(FunctionMenu,ID_Service_Start,"启动服务")
        self.Service_Start.SetBitmap(wx.Bitmap("Picture/service_start.png"))
        self.Service_Start.SetTextColour('black')
        self.Service_Start.SetFont(menu_Font)
        FunctionMenu.Append(self.Service_Start)

        self.Service_Close = wx.MenuItem(FunctionMenu,ID_Service_Close,"关闭服务")
        self.Service_Close.SetBitmap(wx.Bitmap("Picture/service_close.png"))
        self.Service_Close.SetTextColour('black')
        self.Service_Close.SetFont(menu_Font)
        FunctionMenu.Append(self.Service_Close)

        self.ShowTimeInfo= wx.MenuItem(FunctionMenu,ID_ShowTimeInfo,"显示当前打卡时间")
        self.ShowTimeInfo.SetBitmap(wx.Bitmap("Picture/showtimeinfo.png"))
        self.ShowTimeInfo.SetTextColour('black')
        self.ShowTimeInfo.SetFont(menu_Font)
        FunctionMenu.Append(self.ShowTimeInfo)

        self.ShowEmployee = wx.MenuItem(FunctionMenu,ID_ShowEmployee,"显示当前总员工数量以及姓名")
        self.ShowEmployee.SetBitmap(wx.Bitmap("Picture/showemployee.png"))
        self.ShowEmployee.SetTextColour('black')
        self.ShowEmployee.SetFont(menu_Font)
        FunctionMenu.Append(self.ShowEmployee)

        self.SettingTime = wx.MenuItem(FunctionMenu,ID_SettingTime,"上班打卡时间设置")
        self.SettingTime.SetBitmap(wx.Bitmap("Picture/settingtime.png"))
        self.SettingTime.SetTextColour('black')
        self.SettingTime.SetFont(menu_Font)
        FunctionMenu.Append(self.SettingTime)

        self.SettingLeaveTime = wx.MenuItem(FunctionMenu,ID_SettingLeaveTime,"下班打卡时间设置")
        self.SettingLeaveTime.SetBitmap(wx.Bitmap("Picture/settingleavetime.png"))
        self.SettingLeaveTime.SetTextColour('black')
        self.SettingLeaveTime.SetFont(menu_Font)
        FunctionMenu.Append(self.SettingLeaveTime)

        #生成考勤表菜单
        LogMenu = wx.Menu()
        self.SearchLogList = wx.MenuItem(LogMenu,ID_SearchLogList,"搜索考勤表")
        self.SearchLogList.SetBitmap(wx.Bitmap("Picture/openloglist_simple.png"))
        self.SearchLogList.SetTextColour('black')
        self.SearchLogList.SetFont(menu_Font)
        LogMenu.Append(self.SearchLogList)

        self.SearchLogList_Test = wx.MenuItem(LogMenu,ID_SearchLogList_Test,"搜索测试考勤表")
        self.SearchLogList_Test.SetBitmap(wx.Bitmap("Picture/openloglist_complex.png"))
        self.SearchLogList_Test.SetTextColour('black')
        self.SearchLogList_Test.SetFont(menu_Font)
        LogMenu.Append(self.SearchLogList_Test)

        self.CloseLogList = wx.MenuItem(LogMenu,ID_CloseLogList,"关闭考勤表")
        self.CloseLogList.SetBitmap(wx.Bitmap("Picture/closeloglist.png"))
        self.CloseLogList.SetTextColour('black')
        self.CloseLogList.SetFont(menu_Font)
        self.CloseLogList.Enable(False)
        LogMenu.Append(self.CloseLogList)

        self.ExportLogList = wx.MenuItem(LogMenu,ID_ExportLogList,"导出当前考勤表")
        self.ExportLogList.SetBitmap(wx.Bitmap("Picture/exportloglist.png"))
        self.ExportLogList.SetTextColour('black')
        self.ExportLogList.SetFont(menu_Font)
        self.ExportLogList.Enable(False)
        LogMenu.Append(self.ExportLogList)

        self.UpLoad_LogList = wx.MenuItem(LogMenu,ID_UpLoad_LogList,"云端上传考勤表")
        self.UpLoad_LogList.SetBitmap(wx.Bitmap("Picture/upload_loglist.png"))
        self.UpLoad_LogList.SetTextColour('black')
        self.UpLoad_LogList.SetFont(menu_Font)
        LogMenu.Append(self.UpLoad_LogList)

        self.DownLoad_LogList = wx.MenuItem(LogMenu,ID_DownLoad_LogList,"云端下载考勤表")
        self.DownLoad_LogList.SetBitmap(wx.Bitmap("Picture/download_loglist.png"))
        self.DownLoad_LogList.SetTextColour('black')
        self.DownLoad_LogList.SetFont(menu_Font)
        LogMenu.Append(self.DownLoad_LogList)

        #生成测试菜单
        TestMenu = wx.Menu()
        self.Register = wx.MenuItem(TestMenu,ID_Register,"测试新建录入")
        self.Register.SetBitmap(wx.Bitmap("Picture/register.png"))
        self.Register.SetTextColour('black')
        self.Register.SetFont(menu_Font)
        TestMenu.Append(self.Register)

        self.PunchCard = wx.MenuItem(TestMenu,ID_PunchCard,"测试上班打卡")
        self.PunchCard.SetBitmap(wx.Bitmap("Picture/punchcard.png"))
        self.PunchCard.SetTextColour('black')
        self.PunchCard.SetFont(menu_Font)
        TestMenu.Append(self.PunchCard)

        self.Leave = wx.MenuItem(TestMenu,ID_Leave,"测试下班打卡")
        self.Leave.SetBitmap(wx.Bitmap("Picture/leave.png"))
        self.Leave.SetTextColour('black')
        self.Leave.SetFont(menu_Font)
        TestMenu.Append(self.Leave)

        self.AddFaceData = wx.MenuItem(TestMenu,ID_AddFaceData,"模拟添加人脸数据（50人）")
        self.AddFaceData.SetBitmap(wx.Bitmap("Picture/addfacedata.png"))
        self.AddFaceData.SetTextColour('black')
        self.AddFaceData.SetFont(menu_Font)
        TestMenu.Append(self.AddFaceData)

        self.AddPunchCardInfo_Month = wx.MenuItem(TestMenu,ID_AddPunchCardInfo_Month,"模拟添加打卡信息（月）")
        self.AddPunchCardInfo_Month.SetBitmap(wx.Bitmap("Picture/addpunchcardinfo_month.png"))
        self.AddPunchCardInfo_Month.SetTextColour('black')
        self.AddPunchCardInfo_Month.SetFont(menu_Font)
        TestMenu.Append(self.AddPunchCardInfo_Month)
        
        self.AddPunchCardInfo_Year = wx.MenuItem(TestMenu,ID_AddPunchCardInfo_Year,"模拟添加打卡信息（年）")
        self.AddPunchCardInfo_Year.SetBitmap(wx.Bitmap("Picture/addpunchcardinfo_year.png"))
        self.AddPunchCardInfo_Year.SetTextColour('black')
        self.AddPunchCardInfo_Year.SetFont(menu_Font)
        TestMenu.Append(self.AddPunchCardInfo_Year)

        #生成关于菜单
        HelpMenu = wx.Menu()
        self.InitData = wx.MenuItem(HelpMenu,ID_InitData,"格式化本地所有数据")
        self.InitData.SetBitmap(wx.Bitmap("Picture/initdata.png"))
        self.InitData.SetTextColour('black')
        self.InitData.SetFont(menu_Font)
        HelpMenu.Append(self.InitData)

        self.Introduce = wx.MenuItem(HelpMenu,ID_Help,"程序说明")
        self.Introduce.SetBitmap(wx.Bitmap("Picture/introduce.png"))
        self.Introduce.SetTextColour('black')
        self.Introduce.SetFont(menu_Font)
        HelpMenu.Append(self.Introduce)   
        
        self.About = wx.MenuItem(HelpMenu,ID_About,"关于")
        self.About.SetBitmap(wx.Bitmap("Picture/about.png"))
        self.About.SetTextColour('black')
        self.About.SetFont(menu_Font)
        HelpMenu.Append(self.About)

        MenuBar.Append(FunctionMenu,"&功能菜单")
        MenuBar.Append(LogMenu,"&考勤表菜单")
        MenuBar.Append(TestMenu,"&测试菜单")
        MenuBar.Append(HelpMenu,"&帮助")
        self.SetMenuBar(MenuBar)

        #功能菜单事件绑定
        self.Bind(wx.EVT_MENU,self.OnServiceStartClicked,id=ID_Service_Start)
        self.Bind(wx.EVT_MENU,self.OnServiceCloseClicked,id=ID_Service_Close)
        self.Bind(wx.EVT_MENU,self.OnShowTimeInfoClicked,id=ID_ShowTimeInfo)
        self.Bind(wx.EVT_MENU,self.OnShowEmployeeClicked,id=ID_ShowEmployee)
        self.Bind(wx.EVT_MENU,self.OnSettingTimeClicked,id=ID_SettingTime)
        self.Bind(wx.EVT_MENU,self.OnSettingLeaveTimeClicked,id=ID_SettingLeaveTime)

        #考勤表菜单事件绑定
        self.Bind(wx.EVT_MENU,self.OnSearchLogListClicked,id=ID_SearchLogList)
        self.Bind(wx.EVT_MENU,self.OnSearchLogListTestClicked,id=ID_SearchLogList_Test)
        self.Bind(wx.EVT_MENU,self.OnCloseLogListClicked,id=ID_CloseLogList)
        self.Bind(wx.EVT_MENU,self.OnExportLogListClicked,id=ID_ExportLogList)
        self.Bind(wx.EVT_MENU,self.OnUpLoadLogListClicked,id=ID_UpLoad_LogList)
        self.Bind(wx.EVT_MENU,self.OnDownLoadLogListClicked,id=ID_DownLoad_LogList)

        #测试菜单事件绑定
        self.Bind(wx.EVT_MENU,self.OnRegisterClicked,id=ID_Register)
        self.Bind(wx.EVT_MENU,self.OnPunchCardClicked,id=ID_PunchCard)
        self.Bind(wx.EVT_MENU,self.OnLeaveClicked,id=ID_Leave)
        self.Bind(wx.EVT_MENU,self.OnAddFaceDataClicked,id=ID_AddFaceData)
        self.Bind(wx.EVT_MENU,self.OnAddPunchCardInfoMonthClicked,id=ID_AddPunchCardInfo_Month)
        self.Bind(wx.EVT_MENU,self.OnAddPunchCardInfoYearClicked,id=ID_AddPunchCardInfo_Year)

        #帮助菜单事件绑定
        self.Bind(wx.EVT_MENU,self.OnInitDataClicked,id=ID_InitData)  
        self.Bind(wx.EVT_MENU,self.OnIntroduceClicked,id=ID_Help)
        self.Bind(wx.EVT_MENU,self.OnAboutClicked,id=ID_About)

        #在当前界面中心位置显示
        self.Center()

    #按下启动服务对应事件
    def OnServiceStartClicked(self,event):
        wx.MessageBox(message="老婆亲亲可以加快开发速度哦~^v^~", caption="Your baby")
        pass

    #按下关闭服务对应事件
    def OnServiceCloseClicked(self,event):
        wx.MessageBox(message="老婆亲亲可以加快开发速度哦~^v^~", caption="Your baby")
        pass

    #按下显示当前打卡时间
    def OnShowTimeInfoClicked(self,event):
        self.InfoText.AppendText("*当前最早打卡时间为:"+self.MinTime+"*\r\n")
        self.InfoText.AppendText("*当前最迟打卡时间为:"+self.MaxTime+"*\r\n")
        self.InfoText.AppendText("*当前下班打卡时间为:"+self.LeavingTime+"*\r\n")

    #按下显示当前员工总数以及姓名
    def OnShowEmployeeClicked(self,event):
        reload(FaceRecognize)
        self.InfoText.AppendText("*当前已录入的员工数量为{}人*\r\n".format(len(FaceRecognize.features_known_arr)-1))
        self.InfoText.AppendText("姓名:\t\t工号:\r\n")
        for employee_info in os.listdir(path_face_sample):
            self.InfoText.AppendText(employee_info[11:]+"\t\t"+employee_info[0:11]+"\r\n")

    #按下上班打卡时间设置对应事件
    def OnSettingTimeClicked(self,event):
        MinTime = wx.GetTextFromUser(message="请设置最早打卡时间(24小时制)，注意输入格式(时):(分):(秒)", caption="JoyYang提示", default_value="08:00:00", parent=None)
        MaxTime = wx.GetTextFromUser(message="请设置最迟打卡时间(24小时制)，注意输入格式(时):(分):(秒)", caption="JoyYang提示", default_value="09:00:00", parent=None)
        if MinTime == '' and MaxTime == '':
            wx.MessageBox("已取消设置，如要设置打卡时间，请注意输入格式", caption="JoyYang提示")
        elif MinTime == '' and MaxTime != '':
            wx.MessageBox("未输入最早打卡时间，设置失败", caption="JoyYang提示")
        elif MinTime != '' and MaxTime == '':
            wx.MessageBox("未输入最迟打卡时间，设置失败", caption="JoyYang提示")
        else:
            if len(MinTime) != 8 or len(MaxTime) != 8:
                wx.MessageBox("最早或最迟时间输入格式有误，请重新输入", caption="JoyYang警告")
            else:
                if MinTime[2] != ':' or MinTime[5] != ':':
                    wx.MessageBox("最早打卡时间输入为非标准的时间格式，注意冒号的位置，请重新输入", caption="JoyYang警告")
                elif MaxTime[2] != ':' or MaxTime[5] != ':':
                    wx.MessageBox("最迟打卡时间输入为非标准的时间格式，注意冒号的位置，请重新输入", caption="JoyYang警告")
                else:
                    if not self.Is_Int_Number(MinTime[0:2]) or not self.Is_Int_Number(MinTime[3:5]) or not self.Is_Int_Number(MinTime[6:8]):
                        wx.MessageBox("最早打卡时间数据类型输入错误，请输入整型数字设置时间", caption="JoyYang警告")
                    elif not self.Is_Int_Number(MaxTime[0:2]) or not self.Is_Int_Number(MaxTime[3:5]) or not self.Is_Int_Number(MaxTime[6:8]):
                        wx.MessageBox("最迟打卡时间数据类型输入错误，请输入整型数字设置时间", caption="JoyYang警告")
                    else:
                        if MinTime < "00:00:00" and MinTime > "23:59:59":
                            wx.MessageBox("最早打卡时间超出时间设置范围，时间设置范围应为00:00:00~23:59:59", caption="JoyYang警告")
                        elif MinTime < "00:00:00" and MinTime > "23:59:59":
                            wx.MessageBox("最迟打卡时间超出时间设置范围，时间设置范围应为00:00:00~23:59:59", caption="JoyYang警告")
                        elif MinTime > MaxTime:
                            wx.MessageBox("最早打卡时间不能晚于最迟打卡时间", caption="JoyYang警告")
                        else:
                            TimePart = MinTime.split(':')
                            if TimePart[0] < "08":
                                LimitHour = int(TimePart[0][1]) + 2
                                LimitTime = "0{}".format(LimitHour)+":"+TimePart[1]+":"+TimePart[2]
                            elif TimePart[0] == "08" or TimePart[0] == "09":
                                LimitHour = int(TimePart[0][1]) + 2
                                LimitTime = "{}".format(LimitHour)+":"+TimePart[1]+":"+TimePart[2]
                            elif TimePart[0] == "22" or TimePart[0] == "23":
                                LimitHour = (int(TimePart[0]) + 2) % 24
                                LimitTime = "0{}".format(LimitHour)+":"+TimePart[1]+":"+TimePart[2]
                            else:
                                LimitHour = int(TimePart[0]) + 2
                                LimitTime = "{}".format(LimitHour)+":"+TimePart[1]+":"+TimePart[2]                           
                            if MinTime[0:2] == MaxTime[0:2]:
                                wx.MessageBox("最早打卡时间与最迟打卡时间间隔应大于等于1小时", caption="JoyYang警告")
                            elif MaxTime > LimitTime:
                                wx.MessageBox("最早打卡时间与最迟打卡时间间隔应小于等于2小时", caption="JoyYang警告")
                            else:
                                self.MinTime = MinTime
                                self.MaxTime = MaxTime
                                os.remove(path_setting_time)
                                SettingTime = {
                                    'MinTime':self.MinTime,
                                    'MaxTime':self.MaxTime,
                                    'LeavingTime':self.LeavingTime
                                    }
                                json_object = json.dumps(SettingTime)
                                with open(path_setting_time, 'w') as jsonfile:
                                         jsonfile.write(json_object)
                                         jsonfile.close()
                                self.InfoText.AppendText("打卡时间设置成功\r\n")
                                self.InfoText.AppendText("最早打卡时间更新为:"+self.MinTime+"\r\n")
                                self.InfoText.AppendText("最迟打卡时间更新为:"+self.MaxTime+"\r\n")
    
    #按下下班打卡时间设置对应事件
    def OnSettingLeaveTimeClicked(self,event):
        LeavingTime = wx.GetTextFromUser(message="请设置下班打卡时间(24小时制)，注意输入格式(时):(分):(秒)", caption="JoyYang提示", default_value="17:00:00", parent=None)
        if LeavingTime == '':
            wx.MessageBox("已取消设置，如要设置下班打卡时间，请注意输入格式", caption="JoyYang提示")
        else:
            if len(LeavingTime) != 8:
                wx.MessageBox("下班打卡时间输入格式有误，请重新输入", caption="JoyYang警告")
            else:
                if LeavingTime[2] != ':' or LeavingTime[5] != ':':
                    wx.MessageBox("下班打卡时间输入为非标准的时间格式，注意冒号的位置，请重新输入", caption="JoyYang警告")
                else:
                    if not self.Is_Int_Number(LeavingTime[0:2]) or not self.Is_Int_Number(LeavingTime[3:5]) or not self.Is_Int_Number(LeavingTime[6:8]):
                        wx.MessageBox("下班打卡时间数据类型输入错误，请输入整型数字来设置时间", caption="JoyYang警告")
                    else:
                        if LeavingTime < "00:00:00" or LeavingTime > "23:59:59":
                            wx.MessageBox("下班打卡时间超出时间设置范围，时间设置范围应为00:00:00~23:59:59", caption="JoyYang警告")
                        elif LeavingTime < self.MaxTime:
                            wx.MessageBox("下班打卡时间不得早于最迟上班打卡时间", caption="JoyYang警告")
                        else:
                            self.LeavingTime = LeavingTime
                            os.remove(path_setting_time)
                            SettingTime = {
                                'MinTime':self.MinTime,
                                'MaxTime':self.MaxTime,
                                'LeavingTime':self.LeavingTime
                                }
                            json_object = json.dumps(SettingTime)
                            with open(path_setting_time, 'w') as jsonfile:
                                     jsonfile.write(json_object)
                                     jsonfile.close()
                            self.InfoText.AppendText("下班打卡时间设置成功\r\n")
                            self.InfoText.AppendText("下班打卡时间更新为:"+self.LeavingTime+"\r\n")

    #判断字符串中是否为整型数字
    def Is_Int_Number(self,str):
        try:
            int(str)
            return True
        except ValueError:
            return False

    #按下搜索考勤表对应事件
    def OnSearchLogListClicked(self,event):
        self.NLP_Deal(self.database)

    #按下搜索测试考勤表对应事件
    def OnSearchLogListTestClicked(self,event):
        #连接至测试数据库
        database = ApplicationTool.DataBase("./Data/testlist.db")
        self.NLP_Deal(database)

    #检查日期输入是否合法
    def CheckDate(self,date):
        #小月号数
        Small_Month = [4,6,9,11]
        #大月号数
        Big_Month = [1,3,5,7,8,10,12]
        #中月号数
        Middle_Month = [2]
        year = int(date[0])
        month = int(date[1])
        day = int(date[2])
        if month in Small_Month:
            if 1 <= day <= 30:
                return True
            else:
                return False
        elif month in Big_Month:
            if 1 <= day <= 31:
                return True
            else:
                return False
        elif month in Middle_Month:
            #判断不是闰年
            if year % 4 == 0 and year % 100 == 0 and year % 400 != 0:
                if 1 <= day <= 28:
                    return True
                else:
                    return False
            #判断不是闰年
            elif year % 4 != 0:
                if 1<= day <= 28:
                    return True
                else:
                    return False
            #判断是闰年
            else:
                if 1 <= day <= 29:
                    return True
                else:
                    return False
        else:
            return False

    #自然语言处理过程
    def NLP_Deal(self,database):
        parsetool = ApplicationTool.AIParse()
        text = wx.GetTextFromUser(message="请输入搜索内容", caption="JoyYang提示", default_value="查询2020年2月1日至2020年3月31日的打卡记录", parent=None)
        #必要的查询内容关键字
        Key_Word = ["打卡","迟到","早退"]
        Emotion = ["高兴","舒坦","愉悦","闲适","平常","紧张","惆怅","消沉","沮丧","不快"]
        #查询模式标志
        model_flag = 'obscure'
        for employee in os.listdir(path_face_sample):
            if employee[11:] in text:
                ID = employee[0:11]
                text = text.replace(employee[11:],"杨某")
                model_flag = 'specific'
                break
        #同义词手动转译
        text = text.replace("检索","查询")
        text = text.replace("查找","查询")
        text = text.replace("~","至")
        text = text.replace("是","为")
        text = text.replace("到","至",text.count("到")-1)
        if not "查询" in text or not "年" in text or not "月" in text or not "日" in text or not "至" in text or not "的" in text:
            wx.MessageBox("缺少检索必要关键字，请重新输入", caption="JoyYang警告")
        elif text.count("年") != 2 or text.count("月") != 2 or text.count("日") != 2:
            wx.MessageBox("日期条件关键字有误，请重新输入", caption="JoyYang警告")
        elif not text[text.find("的",text.find("至"),-1)+1:text.find("的",text.find("至"),-1)+3] in Key_Word:
            wx.MessageBox("缺少必要查询内容，请重新输入", caption="JoyYang警告")
        else:
            check_flag = parsetool.CheckText(text)
            if check_flag == False:
                wx.MessageBox("无法分析该语句，请检查输入语句逻辑是否有误", caption="JoyYang警告")
            else:
                if self.Is_Int_Number(text[text.find("年",0,text.find("至"))-4:text.find("年",0,text.find("至"))]) == False:
                    wx.MessageBox("起始日期“年”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                elif self.Is_Int_Number(text[text.find("月",0,text.find("至"))-2:text.find("月",0,text.find("至"))]) == False and self.Is_Int_Number(text[text.find("月",0,text.find("至"))-1:text.find("月",0,text.find("至"))]) == False:
                    wx.MessageBox("起始日期“月”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                elif self.Is_Int_Number(text[text.find("日",0,text.find("至"))-2:text.find("日",0,text.find("至"))]) == False and self.Is_Int_Number(text[text.find("日",0,text.find("至"))-1:text.find("日",0,text.find("至"))]) == False:
                    wx.MessageBox("起始日期“日”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                elif self.Is_Int_Number(text[text.find("年",text.find("至"),-1)-4:text.find("年",text.find("至"),-1)]) == False:
                    wx.MessageBox("终止日期“年”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                elif self.Is_Int_Number(text[text.find("月",text.find("至"),-1)-2:text.find("月",text.find("至"),-1)]) == False and self.Is_Int_Number(text[text.find("月",text.find("至"),-1)-1:text.find("月",text.find("至"),-1)]) == False:
                    wx.MessageBox("终止日期“月”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                elif self.Is_Int_Number(text[text.find("日",text.find("至"),-1)-2:text.find("日",text.find("至"),-1)]) == False and self.Is_Int_Number(text[text.find("日",text.find("至"),-1)-1:text.find("日",text.find("至"),-1)]) == False:
                    wx.MessageBox("终止日期“日”条件有误，请参考程序说明的提示", caption="JoyYang警告")
                else:
                    start_date = []
                    end_date = []
                    start_date.append(text[text.find("年",0,text.find("至"))-4:text.find("年",0,text.find("至"))])
                    if self.Is_Int_Number(text[text.find("月",0,text.find("至"))-2:text.find("月",0,text.find("至"))]) == False and self.Is_Int_Number(text[text.find("月",0,text.find("至"))-1:text.find("月",0,text.find("至"))]) == True:
                        start_date.append(text[text.find("月",0,text.find("至"))-1:text.find("月",0,text.find("至"))])
                    else:
                        start_date.append(text[text.find("月",0,text.find("至"))-2:text.find("月",0,text.find("至"))])
                    if self.Is_Int_Number(text[text.find("日",0,text.find("至"))-2:text.find("日",0,text.find("至"))]) == False and self.Is_Int_Number(text[text.find("日",0,text.find("至"))-1:text.find("日",0,text.find("至"))]) == True:
                        start_date.append(text[text.find("日",0,text.find("至"))-1:text.find("日",0,text.find("至"))])
                    else:
                        start_date.append(text[text.find("日",0,text.find("至"))-2:text.find("日",0,text.find("至"))])
                    end_date.append(text[text.find("年",text.find("至"),-1)-4:text.find("年",text.find("至"),-1)])
                    if self.Is_Int_Number(text[text.find("月",text.find("至"),-1)-2:text.find("月",text.find("至"),-1)]) == False and self.Is_Int_Number(text[text.find("月",text.find("至"),-1)-1:text.find("月",text.find("至"),-1)]) == True:
                        end_date.append(text[text.find("月",text.find("至"),-1)-1:text.find("月",text.find("至"),-1)])
                    else:
                        end_date.append(text[text.find("月",text.find("至"),-1)-2:text.find("月",text.find("至"),-1)])
                    if self.Is_Int_Number(text[text.find("日",text.find("至"),-1)-2:text.find("日",text.find("至"),-1)]) == False and self.Is_Int_Number(text[text.find("日",text.find("至"),-1)-1:text.find("日",text.find("至"),-1)]) == True:
                        end_date.append(text[text.find("日",text.find("至"),-1)-1:text.find("日",text.find("至"),-1)])
                    else:
                        end_date.append(text[text.find("日",text.find("至"),-1)-2:text.find("日",text.find("至"),-1)])
                    if self.CheckDate(start_date) == False or self.CheckDate(end_date) == False:
                        wx.MessageBox("日期数据类型输入正确，但不合法，请检查", caption="JoyYang警告")
                    else:
                        content = text[text.find("的",text.find("至"),-1)+1:text.find("的",text.find("至"),-1)+3]
                        if "情绪" in text:
                            if not "为" in text:
                                wx.MessageBox("含有情绪查询的条件关键字有误，请参考程序说明", caption="JoyYang警告")
                            elif not text[text.find("为",text.find("至"),-1)+1:text.find("为",text.find("至"),-1)+3] in Emotion:
                                wx.MessageBox("含有情绪查询的条件关键字有误，请参考程序说明", caption="JoyYang警告")
                            else:
                                emotion = text[text.find("为",text.find("至"),-1)+1:text.find("为",text.find("至"),-1)+3]
                                if model_flag == 'obscure':
                                    LogData = database.Obscure_Search(start_date,end_date,emotion,content)
                                    self.DisPlaySearchResult(LogData)
                                else:
                                    LogData = database.Specific_Search(start_date,end_date,ID,emotion,content)
                                    self.DisPlaySearchResult(LogData)
                        else:
                            emotion = "None"
                            if model_flag == 'obscure':
                                LogData = database.Obscure_Search(start_date,end_date,emotion,content)
                                self.DisPlaySearchResult(LogData)
                            else:
                                LogData = database.Specific_Search(start_date,end_date,ID,emotion,content)
                                self.DisPlaySearchResult(LogData)

    #展示查询数据方法
    def DisPlaySearchResult(self,LogData):
        self.SearchLogList.Enable(False)
        self.SearchLogList_Test.Enable(False)
        self.CloseLogList.Enable(True)
        wx.Frame.SetSize(self,size=(1220,660))
        log_ID = LogData.get('ID')
        log_name = LogData.get('Name')
        log_date = LogData.get('Date')
        log_punchcard = LogData.get('PunchCard')
        log_late = LogData.get('Late')
        log_latetime = LogData.get('LateTime')
        log_leave = LogData.get('Leave')
        log_leavetime = LogData.get('LeaveTime')
        log_worktime = LogData.get('WorkTime')
        log_workstatus = LogData.get('WorkStatus')
        self.grid = wx.grid.Grid(self,pos=(300,0),size=(900,600))
        self.grid.CreateGrid(len(log_ID), 10)
        for i in range(len(log_ID)):
            for j in range(10):
                self.grid.SetCellAlignment(i,j,wx.ALIGN_CENTER,wx.ALIGN_CENTER)
        self.grid.SetColLabelValue(0, "工号")
        self.grid.SetColLabelValue(1, "姓名")
        self.grid.SetColLabelValue(2, "日期")
        self.grid.SetColLabelValue(3, "上班打卡时间")        
        self.grid.SetColLabelValue(4, "出勤情况")
        self.grid.SetColLabelValue(5, "迟到时长")
        self.grid.SetColLabelValue(6, "早退情况")
        self.grid.SetColLabelValue(7, "下班打卡时间")
        self.grid.SetColLabelValue(8, "工作时长")
        self.grid.SetColLabelValue(9, "工作状态")
        self.grid.SetColSize(0,100)
        self.grid.SetColSize(1,60)
        self.grid.SetColSize(2,80)
        self.grid.SetColSize(3,100)
        self.grid.SetColSize(4,60)
        self.grid.SetColSize(5,80)
        self.grid.SetColSize(6,60)
        self.grid.SetColSize(7,100)
        self.grid.SetColSize(8,80)
        self.grid.SetColSize(9,80)
        self.create_excellist()
        for i in range(len(log_ID)):
            row_data = []
            if log_name[i] == "杨舒粤":
                self.grid.SetCellTextColour('pink')
            else:
                self.grid.SetCellTextColour('blue')
            self.grid.SetCellTextColour(i,0,'purple')
            self.grid.SetCellValue(i,0,log_ID[i])
            row_data.append(log_ID[i])
            self.grid.SetCellValue(i,1,log_name[i])
            row_data.append(log_name[i])
            self.grid.SetCellValue(i,2,log_date[i])
            row_data.append(log_date[i])
            self.grid.SetCellValue(i,3,log_punchcard[i])
            row_data.append(log_punchcard[i])
            if log_late[i] == "迟到":
                self.grid.SetCellTextColour(i,4,'red')
                self.grid.SetCellTextColour(i,5,'red')
                self.grid.SetCellValue(i,4,log_late[i])
                self.grid.SetCellValue(i,5,log_latetime[i])
            elif log_late[i] == "未迟到":
                self.grid.SetCellTextColour(i,4,'green')
                self.grid.SetCellTextColour(i,5,'green')
                self.grid.SetCellValue(i,4,log_late[i])
                self.grid.SetCellValue(i,5,log_latetime[i])
            else:
                self.grid.SetCellTextColour(i,4,'yellow')
                self.grid.SetCellTextColour(i,5,'yellow')
                self.grid.SetCellValue(i,4,log_late[i])
                self.grid.SetCellValue(i,5,log_latetime[i])
            row_data.append(log_late[i])
            row_data.append(log_latetime[i])
            if log_leave[i] == "早退":
                self.grid.SetCellTextColour(i,6,'red')
                self.grid.SetCellTextColour(i,7,'red')
                self.grid.SetCellValue(i,6,log_leave[i])
                self.grid.SetCellValue(i,7,log_leavetime[i])
            elif log_leave[i] == "未早退":
                self.grid.SetCellTextColour(i,6,'green')
                self.grid.SetCellTextColour(i,7,'green')
                self.grid.SetCellValue(i,6,log_leave[i])
                self.grid.SetCellValue(i,7,log_leavetime[i])
            else:
                self.grid.SetCellTextColour(i,6,'yellow')
                self.grid.SetCellTextColour(i,7,'yellow')
                self.grid.SetCellValue(i,6,log_leave[i])
                self.grid.SetCellValue(i,7,log_leavetime[i])
            row_data.append(log_leave[i])
            row_data.append(log_leavetime[i])
            self.grid.SetCellValue(i,8,log_worktime[i])
            row_data.append(log_worktime[i])
            self.grid.SetCellValue(i,9,log_workstatus[i])
            row_data.append(log_workstatus[i])
            self.insert_a_row("A{}".format(i+2),row_data)
        self.save_as_excellist()
        self.ExportLogList.Enable(True)

    #按下关闭考勤表对应事件
    def OnCloseLogListClicked(self,event):
        self.SearchLogList.Enable(True)
        self.SearchLogList_Test.Enable(True)
        self.CloseLogList.Enable(False)
        self.ExportLogList.Enable(False)
        wx.Frame.SetSize(self,size=(1200,660))
        self.grid.Destroy()
        self.CreateGallery()
        if os.path.exists("./Data/log_list.xlsx"):
            os.remove("./Data/log_list.xlsx")

    #按下导出当前考勤表对应事件
    def OnExportLogListClicked(self,event):
        self.ExportLogList.Enable(False)
        shutil.copyfile("./Data/log_list.xlsx","./Export/log_list.xlsx")
        gaugeframe = ApplicationTool.SimpleGauge("导出中，请稍后...",self.InfoText)
        gaugeframe.Show()
        os.remove("./Data/log_list.xlsx")

    #导出当前考勤表方法
    def create_excellist(self):
        header = ["工号","姓名","打卡日期","上班打卡时间","出勤情况","迟到时长","早退情况","下班打卡时间","工作时长","工作状态"]
        self.workbook = xlsxwriter.Workbook("./Data/log_list.xlsx")
        self.workfomat = self.workbook.add_format()
        self.workfomat.set_align('center')
        headformat = self.workbook.add_format()
        headformat.set_bold(1)
        headformat.set_align('center')
        self.worksheet = self.workbook.add_worksheet('PunchCardList')
        self.worksheet.write_row("A1",header,headformat)
        self.worksheet.set_column('A:J',15)
    def insert_a_row(self,index,insert_row):
        self.worksheet.write_row(index,insert_row,self.workfomat)
    def save_as_excellist(self):
        self.workbook.close()

    #按下云端保存考勤表对应事件
    def OnUpLoadLogListClicked(self,event):
        Transport = ApplicationTool.ConnectServer(IP,Port)
        Transport.Connect()
        Transport.UpLoadFile(path_database,path_remote+path_database.split('/')[-1])
        Transport.UpLoadFile(path_feature_all,path_remote+path_feature_all.split('/')[-1])
        Transport.UpLoadFile(path_setting_time,path_remote+path_setting_time.split('/')[-1])
        gaugeframe = ApplicationTool.SimpleGauge("上传中，请稍后...",self.InfoText)
        gaugeframe.Show()
        Transport.CloseConnect()

    #按下云端下载考勤表对应事件
    def OnDownLoadLogListClicked(self,event):
        Transport = ApplicationTool.ConnectServer(IP,Port)
        Transport.Connect()
        Transport.DownLoadFile(path_remote+path_database.split('/')[-1],path_database)
        Transport.DownLoadFile(path_remote+path_feature_all.split('/')[-1],path_feature_all)
        Transport.DownLoadFile(path_remote+path_setting_time.split('/')[-1],path_setting_time)
        gaugeframe = ApplicationTool.SimpleGauge("下载中，请稍后...",self.InfoText)
        gaugeframe.Show()
        Transport.CloseConnect()

    #按下本机测试新建录入对应事件
    def OnRegisterClicked(self,event):       
        self.Register.Enable(False)
        num_str = ''.join(str(random.choice(range(10))) for i in range(11))
        self.ID = wx.GetTextFromUser(message="输入您的11位工号", caption="JoyYang提示", default_value=num_str, parent=None)
        self.name = wx.GetTextFromUser(message="请输入您的姓名", caption="JoyYang提示", default_value="", parent=None)
        if len(self.ID) != 11:
            wx.MessageBox(message="工号非11位，请重新输入", caption="JoyYang警告")
            self.Register.Enable(True)
            self.ID = ""
            self.name = ""
        else:
            self.name_info = self.ID + self.name
            #监测是否重名
            for exsit_name in os.listdir(path_face_sample):
                if self.name_info[0:11] == exsit_name[0:11] or self.name_info[11:] == exsit_name[11:]:               
                    self.name_info = '***'
                    break
            if self.name_info != '' and self.name_info != '***':
                os.makedirs(path_face_sample+self.name_info)
                self.InfoText.AppendText("等待捕获人脸样本......\r\n")
                self.OpenCamera("FaceDetector")
            elif self.name_info == '***':
                wx.MessageBox(message="工号或姓名已存在，输入失败", caption="JoyYang警告")
                self.Register.Enable(True)
            else:
                wx.MessageBox(message="输入为空或直接关闭，输入失败", caption="JoyYang警告")
                self.Register.Enable(True)

    #打开摄像头方法
    def OpenCamera(self,window_name):
        num = 0
        cv2.namedWindow(window_name)  
        capture = cv2.VideoCapture(0)
        classfier = cv2.CascadeClassifier(path_classifiers)

        while capture.isOpened():
            success, frame = capture.read()
            #摄像头异常时，直接跳出
            if not success:
                capture.release()
                cv2.destroyAllWindows()
                for i in list(range(0,4,1)):
                    cv2.waitKey(1)
                self.InfoText.AppendText("摄像头异常，请重试\r\n")
                break
            self.count += 1
            #将当前帧的图像转换成灰度图像以便于提高计算效率
            grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            #人脸检测，1.2和3分别为图片缩放比例和需要检测的有效点数
            faceRects = classfier.detectMultiScale(grey, scaleFactor=1.2, minNeighbors=3, minSize=(64, 64))
            #大于0则检测到人脸
            if len(faceRects) > 0:    
                #框出视频流中出现的每一张人脸
                for faceRect in faceRects:  
                    x, y, w, h = faceRect               
                    img_name = "%s/%d.jpg"%(path_face_sample+self.name_info, num)
                    image = frame[y-10:y+h+10, x-10:x+w+10]
                    cv2.imencode('.jpg', image)[1].tofile(img_name)
                    cv2.rectangle(frame, (x-10, y-10), (x+w+10, y+h+10), (0,0,0), 2)
                    if num == 0:
                        cv2.imwrite("Data/temp.jpg", image)
                    elif num > 10:
                        break
                    num += 1
                    cv2.putText(frame,"Number:%d" % (num),(x+30, y+30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0,0,0),4)                          
            cv2.imshow(window_name, frame)        
            cv2.waitKey(10)

            #人脸样本采集达到上限则退出
            if num > 10:
                capture.release()
                cv2.destroyAllWindows()
                for i in list(range(0,4,1)):
                    cv2.waitKey(1)
                break

            #计时器达到上限则退出
            if self.count == 30:
                self.InfoText.AppendText("未识别到人脸\r\n")
                capture.release()
                cv2.destroyAllWindows()
                for i in list(range(0,4,1)):
                    cv2.waitKey(1)
                break
        self.DealRegisterEvent(num)

    #人脸录入处理事件
    def DealRegisterEvent(self,num):
        reload(FaceRecognize)
        save_num = num
        if os.path.exists("Data/temp.jpg"):
            img_read = cv2.imread("Data/temp.jpg")
            rects = FaceRecognize.detector(img_read, 1)
            cv2.waitKey(10)
            features_img = FaceRecognize.classifiers.compute_face_descriptor(img_read, FaceRecognize.predictor(img_read, rects[0]))
            for i in range(len(FaceRecognize.features_known_arr)):
                #将此人脸数据与存储的所有人脸数据进行比对
                compare = FaceRecognize.return_euclidean_distance(features_img, FaceRecognize.features_known_arr[i][0:-2])
                #找到了相似脸，判断为重复录入人脸
                if compare == "same": 
                    face_name = FaceRecognize.features_known_arr[i][-1]
                    wx.MessageBox(message="亲爱的"+face_name+"，您已录过人脸了哟", caption="JoyYang警告")
                    self.repeat_register_flag = True
                    break
        
        if self.repeat_register_flag == True:
            shutil.rmtree(path_face_sample+self.name_info)
            self.InfoText.AppendText("重复录入，已删除姓名文件夹\r\n")

        elif save_num < 10 and self.repeat_register_flag == False:
            shutil.rmtree(path_face_sample+self.name_info)
            self.InfoText.AppendText("人脸样本不完整，请重试\r\n")

        else:
            pics = os.listdir(path_face_sample+self.name_info)
            feature_list = []
            feature_average = []
            for i in range(len(pics)):
                pic_path = path_face_sample + self.name_info + "/" + pics[i]
                self.InfoText.AppendText("正在读的人脸图像"+pic_path+"\r\n")
                img = iio.imread(pic_path)
                #将人脸样本图像转换成灰度图像以便于提高计算效率
                img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                dets = FaceRecognize.detector(img_gray, 1)
                if len(dets) != 0:
                    shape = FaceRecognize.predictor(img_gray, dets[0])
                    face_descriptor = FaceRecognize.classifiers.compute_face_descriptor(img_gray, shape)
                    feature_list.append(face_descriptor)
                else:
                    face_descriptor = 0
                    self.InfoText.AppendText("未在人脸样本中识别到的人脸\r\n")

            if len(feature_list)>0:
                #识别到128个特征值可表示一张特定的人脸
                for i in range(128):
                    feature_average.append(0)
                    for j in range(len(feature_list)):
                        feature_average[i] += feature_list[j][i]
                    feature_average[i] = (feature_average[i])/len(feature_list)
                feature_average.append(self.ID)
                feature_average.append(self.name)

                with open(path_feature_all, 'a+', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(feature_average)
                    self.InfoText.AppendText("保存成功，人脸数据已写入人脸库\r\n")

                self.aispeech.SendText("亲爱的"+self.name+"您已经成功录入人脸，祝您工作愉快")
                self.aispeech.Play()

        #初始化各项参数
        self.name = ""
        self.count = 0
        self.Register.Enable(True)
        self.register_flag = False
        if os.path.exists("Data/temp.jpg"):
            os.remove("Data/temp.jpg")
    
    #按下测试上班打卡对应事件
    def OnPunchCardClicked(self,event):
        PunchCardFrame = FaceRecognize.PunchcardUI(None,self.InfoText,self.MinTime,self.MaxTime,self.LeavingTime,"punchcard")
        PunchCardFrame.Show()

    #按下测试下班打卡对应事件
    def OnLeaveClicked(self,event):
        PunchCardFrame = FaceRecognize.PunchcardUI(None,self.InfoText,self.MinTime,self.MaxTime,self.LeavingTime,"leave")
        PunchCardFrame.Show()
    
    #按下模拟添加人脸数据对应事件
    def OnAddFaceDataClicked(self,event):
        reload(FaceRecognize)
        if len(FaceRecognize.features_known_arr)-1 >= 50:
            wx.MessageBox(message="已经进行过模拟添加人脸数据",caption="JoyYang警告")
            self.InfoText.AppendText("检测到重复的操作，添加失败\r\n")
        else:
            for i in range(50):
                Simulate_Data = []
                if i < 10:
                    WorkNum = "1631032040{}".format(i)
                else:
                    WorkNum = "163103204{}".format(i)
                FirstName = "赵孙李周吴郑王冯陈卫蒋沈韩杨朱秦许何施张孔曹严华陶穆姚甄段梁"
                LastName = ["明哲","涵煦","品鸿","潇然","俊驰","浩宇","泽洋","弘文","鹏涛","寻欢","雨泽","泽峻","书昀","冠博","绍远",
                            "若霞","水静","涵月","悠柔","楚馨","子墨","慧语","紫瑶","若尘","希月","雨欣","佳雪","颖琳","诗芮","丽文",
                            "哲","皓","景","恒","晨","志","奕","俊","晋","磊",
                            "婵","沐","玥","欣","曦","姗","瑜","琳","婷","琪"]
                first_name = random.choice(FirstName)
                last_name = random.choice(LastName)
                for j in range(128):
                    Simulate_Data.append(0)
                Simulate_Data.append(WorkNum)
                Simulate_Data.append(first_name+last_name)
                with open(path_feature_all, 'a+', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(Simulate_Data)   
                name_info = WorkNum+first_name+last_name
                os.makedirs(path_face_sample+name_info)
            self.InfoText.AppendText("模拟人脸数据已全部写入数据库\r\n")

    #按下模拟添加打卡数据（月）对应事件
    def OnAddPunchCardInfoMonthClicked(self,event):
        #连接至测试数据库
        datebase = ApplicationTool.DataBase("./Data/testlist.db")
        text = wx.GetTextFromUser(message="请输入生成某年某月的数据", caption="JoyYang提示", default_value="", parent=None)
        if not "年" in text or not "月" in text:
            wx.MessageBox(message="缺少关键字“年”或“月”",caption="JoyYang警告")
        elif len(text) < 7 or len(text) > 8 or self.Is_Int_Number(text[:text.find("年")]) == False or self.Is_Int_Number(text[text.find("年")+1:text.find("月")]) == False:
            wx.MessageBox(message="输入不合法，请参考程序说明",caption="JoyYang警告")
        else:
            frame = ApplicationTool.ComplexGauge(text,os.listdir(path_face_sample),self.MinTime,self.MaxTime,self.LeavingTime,datebase,self.InfoText)
            frame.Show()

    #按下模拟添加打卡数据（年）对应事件
    def OnAddPunchCardInfoYearClicked(self,event):
        #连接至测试数据库
        datebase = ApplicationTool.DataBase("./Data/testlist.db")
        text = wx.GetTextFromUser(message="请输入生成某年的数据", caption="JoyYang提示", default_value="", parent=None)
        if not "年" in text:
            wx.MessageBox(message="缺少关键字“年”",caption="JoyYang警告")
        elif len(text) != 5 or self.Is_Int_Number(text[:text.find("年")]) == False:
            wx.MessageBox(message="输入不合法，请参考程序说明",caption="JoyYang警告")
        else:   
            frame = ApplicationTool.ComplexGauge(text,os.listdir(path_face_sample),self.MinTime,self.MaxTime,self.LeavingTime,datebase,self.InfoText)
            frame.Show()

    #按下格式化本地所有数据对应事件
    def OnInitDataClicked(self,event):
        prompt = "本地所有打卡数据与人脸库数据都将被清除，该操作不可逆，是否继续？"
        dig = wx.MessageDialog(None,prompt,"JoyYang提示",wx.YES_NO|wx.ICON_QUESTION)
        if dig.ShowModal() == wx.ID_YES:
            shutil.rmtree("Data/face_img_database")
            shutil.rmtree("./Export")
            os.remove("Data/feature_all.csv")        
            os.remove("Data/loglist.db")
            os.remove("Data/setting_time.json")
            os.makedirs("Data/face_img_database")
            os.makedirs("./Export")
            if os.path.exists("Speech.mp3"):
                os.remove("Speech.mp3")
            if os.path.exists("Speech.wav"):
                os.remove("Speech.wav")
            SettingTime = {
                'MinTime':"07:00:00",
                'MaxTime':"09:00:00",
                'LeavingTime':"17:00:00"
            }
            json_object = json.dumps(SettingTime)
            with open(path_setting_time, 'w') as jsonfile:
                     jsonfile.write(json_object)
                     jsonfile.close()

            with open(path_feature_all, 'w', newline='') as csvfile:
                 csv_write = csv.writer(csvfile)
                 csv_head = []
                 for i in range(128):
                     csv_head.append(0)
                 csv_head.append("ID")
                 csv_head.append("Name")            
                 csv_write.writerow(csv_head)

            self.database = ApplicationTool.DataBase(path_database)
            database = ApplicationTool.DataBase("./Data/testlist.db")
            gaugeframe = ApplicationTool.SimpleGauge("重置中，请稍后...",self.InfoText)
            gaugeframe.Show()
        else:
            wx.MessageBox("已撤回格式化本地所有数据操作", caption="JoyYang提示")

    #按下程序说明对应事件     
    def OnIntroduceClicked(self,event):
        string = '''程序使用注意事项：\n
        1.添加模拟打卡数据之前应确保已经模拟添加人脸数据。\n
        2.查询是利用关键字以及自然语言处理混合处理的，换句话说，光有关键字没有一个通顺的语句也是不行的，一句话中必须含有“查询”，“至”，“年”，“月”，“日”以及{“打卡”、“迟到”、“早退”}中任意一个，具体可参照默认示例语句。\n
        3.查询分为模糊查询和精准查询，模糊查询只需要包含第2点中提到的关键字组成一句话即可，而精准查询则需要在日期前添加姓名或者在查询内容前添加情绪关键字，例如：查询XXX在2020年2月1日至2020年3月31日的迟到记录、查询2020年2月1日至2020年3月31日情绪为XX的记录\n
        4.查询之后展示区会显示一张表格，可在考勤表菜单栏中点击导出当前考勤表导出至项目目录中Export文件夹下。（注意：查询或导出考勤表后要记得点击考勤表菜单栏下关闭考勤表才会删除数据缓存，释放空间）\n
        5.打卡时间设置中，最早打卡时间与最晚打卡时间的间隔应大于等于1小时，小于等于2小时
        6.除此之外其他操作按照提示即可，若有疑问，请及时联系程序开发者your lovely baby
        '''
        wx.MessageBox(message=string,caption="程序说明")
    
    #按下关于对应事件
    def OnAboutClicked(self,event):
        wx.MessageBox(message="程序制作人:JoyYang\n班级:\n学号:",caption="关于")

    #创建文本信息栏方法
    def CreateInfoText(self):
        self.InfoText = wx.TextCtrl(parent=self,size=(300,600),style=(wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY))
        #字体颜色设置
        self.InfoText.SetForegroundColour('black')
        #字体参数设置
        font = wx.Font()
        font.SetPointSize(10)
        font.SetFamily(wx.ROMAN)
        self.InfoText.SetFont(font)
        self.InfoText.SetBackgroundColour('white')
        self.InfoText.AppendText("欢迎进入打卡系统总控制台\r\n等待操作中...\r\n")

    #创建展示栏方法
    def CreateGallery(self):
        self.pic = wx.Image(MainPic, wx.BITMAP_TYPE_ANY).Scale(880, 600)
        self.bmp = wx.StaticBitmap(parent=self, pos=(300,0), bitmap=wx.Bitmap(self.pic))   

class MainForm(wx.App):
    def OnInit(self):
        self.frame = MainUI(None)
        self.frame.Show()
        return True

mainForm = MainForm()
mainForm.MainLoop()