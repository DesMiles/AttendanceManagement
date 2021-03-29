import os
import wx
import sqlite3
import datetime
import paramiko
import wave
import pyaudio
import time
import random
from aip import AipSpeech
from aip import AipNlp

#计算并返回两个时间点间的时长（不计算秒位）
def return_time_sustained(time_former,time_latter):
    #整理时位数据
    if time_former[0] == "0":
        hour_former = int(time_former[1])
    else:
        hour_former = int(time_former[0:2])
    if time_latter[0] == "0":
        hour_latter = int(time_latter[1])
    else:
        hour_latter = int(time_latter[0:2])
    #整理分位数据
    if time_former[3] == "0":
        minute_former = int(time_former[4])
    else:
        minute_former = int(time_former[3:5])
    if time_latter[3] == "0":
        minute_latter = int(time_latter[4])
    else:
        minute_latter = int(time_latter[3:5])

    minute = minute_latter - minute_former
    if minute < 0:
        hour_latter -= 1
        minute += 60
    hour = hour_latter - hour_former
    sust = "{}小时{}分钟".format(hour,minute)
    return sust

#数据库
class DataBase:
    def __init__(self,path):
        LocalTime = datetime.datetime.now()
        self.path = path        
        self.Table = "LogInfo_"+str(LocalTime.year)+"_"+str(LocalTime.month)
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()

        SQL = '''CREATE TABLE IF NOT EXISTS {}(
        ID TEXT NOT NULL,                   /*员工工号*/
        Name TEXT NOT NULL,                 /*员工姓名*/
        Date_Info TEXT NOT NULL,            /*员工打卡日期*/
        PunchCard TEXT NOT NULL,            /*员工上班打卡时间*/
        Late TEXT NOT NULL,                 /*员工当日出勤情况*/
        LateTime TEXT NOT NULL,             /*员工当日迟到时长*/
        Leave TEXT NOT NULL,                /*员工当日早退情况*/
        LeaveTime TEXT NOT NULL,            /*员工下班打卡时间*/
        WorkTime TEXT NOT NULL,             /*员工当日工作时长*/
        WorkStatus TEXT NOT NULL            /*员工工作状态*/
        )'''.format(self.Table)

        cursor.execute(SQL)
        cursor.close()
        connect.commit()
        connect.close()
    
    #插入员工工号、员工姓名、员工打卡日期、员工当日出勤情况、员工当日迟到时长等信息
    def Insert(self,list):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        SQL = "INSERT INTO {} (ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus) VALUES (?,?,?,?,?,?,?,?,?,?)".format(self.Table)
        cursor.execute(SQL,(list[0],list[1],list[2],list[3],list[4],list[5],list[6],list[7],list[8],list[9]))
        cursor.close()
        connect.commit()
        connect.close()
    
    #模拟插入数据调用方法
    def Insert_Data(self,table,list):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        not_exist_flag = True
        SQL = "SELECT * FROM {}".format(table)
        try:
            cursor.execute(SQL)
        except:
            not_exist_flag = False
        if not_exist_flag == True:
            SQL = '''CREATE TABLE {}(
            ID TEXT NOT NULL,                   /*员工工号*/
            Name TEXT NOT NULL,                 /*员工姓名*/
            Date_Info TEXT NOT NULL,            /*员工打卡日期*/
            PunchCard TEXT NOT NULL,            /*员工上班打卡时间*/
            Late TEXT NOT NULL,                 /*员工当日出勤情况*/
            LateTime TEXT NOT NULL,             /*员工当日迟到时长*/
            Leave TEXT NOT NULL,                /*员工当日早退情况*/
            LeaveTime TEXT NOT NULL,            /*员工下班打卡时间*/
            WorkTime TEXT NOT NULL,             /*员工当日工作时长*/
            WorkStatus TEXT NOT NULL            /*员工工作状态*/
            )'''.format(table)
            cursor.execute(SQL)
            connect.commit()        
            SQL = "INSERT INTO {} (ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus) VALUES (?,?,?,?,?,?,?,?,?,?)".format(table)
            cursor.execute(SQL,(list[0],list[1],list[2],list[3],list[4],list[5],list[6],list[7],list[8],list[9]))
            cursor.close()
            connect.commit()
            connect.close()
        else:
            print("表"+table+"已存在")
            cursor.close()
            connect.close()

    #修改早退情况、下班打卡时间、工作时长等信息
    def ModifyLeaveInfo(self,id,date,leave_info):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        SQL = "UPDATE {} SET Leave=?,LeaveTime=?,WorkTime=? WHERE ID=? AND Date_Info=?".format(self.Table)
        cursor.execute(SQL,(leave_info[0],leave_info[1],leave_info[2],id,date))
        cursor.close()
        connect.commit()
        connect.close()
    
    #查询上班打卡时间
    def FindPunchCardTime(self,id,date):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        SQL = "SELECT PunchCard FROM {} WHERE ID=? AND Date_Info=?".format(self.Table)
        cursor.execute(SQL,(id,date))
        get_data = cursor.fetchone()
        cursor.close()
        connect.close()
        return get_data[0]

    #查询出勤情况信息
    def FindPunchCardInfo(self,id,date):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        SQL = "SELECT Late FROM {} WHERE ID=? AND Date_Info=?".format(self.Table)
        cursor.execute(SQL,(id,date))
        get_data = cursor.fetchone()
        cursor.close()
        connect.close()
        try:
            print(get_data[0])
            return True
        except TypeError:
            return False

    #查询早退情况信息
    def FindLeaveInfo(self,id,date):
        connect = sqlite3.connect(self.path)
        cursor = connect.cursor()
        SQL = "SELECT Leave FROM {} WHERE ID=? AND Date_Info=?".format(self.Table)
        cursor.execute(SQL,(id,date))
        get_data = cursor.fetchone()
        cursor.close()
        connect.close()
        if get_data[0] == "Unkown":
            return False
        else:
            return True
    
    #精准查询
    def Specific_Search(self,start_date,end_date,id,emotion,content):
        #小月号数
        Small_Month = [4,6,9,11]
        #大月号数
        Big_Month = [1,3,5,7,8,10,12]
        #中月号数
        Middle_Month = [2]
        start_year = int(start_date[0])
        start_month = int(start_date[1])
        start_day = int(start_date[2])
        deadline = "{}-{}-{}".format(end_date[0],end_date[1],end_date[2])
        connect = sqlite3.connect(self.path)       
        log_ID = []
        log_name = []
        log_date = []
        log_punchcard = []
        log_late = []
        log_latetime = []
        log_leave = []
        log_leavetime = []
        log_worktime = []
        log_workstatus = []
        while True:
            cursor = connect.cursor()
            table = "LogInfo_"+"{}".format(start_year)+"_"+"{}".format(start_month)
            if emotion == "None":
                if content == "迟到":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Late=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content,id))
                elif content == "早退":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Leave=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content,id))
                else:
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),id))
            else:
                if content == "迟到":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Late=? AND WorkStatus=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content,emotion,id))
                elif content == "早退":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Leave=? AND WorkStatus=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),contentn,emotion,id))
                else:
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND WorkStatus=? AND ID=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),emotion,id))
            all_data = cursor.fetchall()
            cursor.close()
            for row in all_data:
                log_ID.append(row[0])
                log_name.append(row[1])
                log_date.append(row[2])
                log_punchcard.append(row[3])
                log_late.append(row[4])
                log_latetime.append(row[5])
                log_leave.append(row[6])
                log_leavetime.append(row[7])
                log_worktime.append(row[8])
                log_workstatus.append(row[9])
            if "{}-{}-{}".format(start_year,start_month,start_day) == deadline:
                break
            start_day += 1
            if start_month in Small_Month and start_day == 31:
                start_day = 1
                start_month += 1
            elif start_month in Big_Month and start_day == 32:
                start_day = 1
                start_month += 1
            elif start_month in Middle_Month and start_day == 29:
                #判断不是闰年
                if start_year % 4 == 0 and start_year % 100 == 0 and start_year % 400 != 0:
                    start_day = 1
                    start_month += 1
                #判断不是闰年
                elif start_year % 4 != 0:
                    start_day = 1
                    start_month += 1
                #判断是闰年
                else:
                    pass
            #闰年间二月的情况处理
            elif start_month in Middle_Month and start_day == 30:
                start_day = 1
                start_month += 1
            if start_month == 13:
                start_month = 1
                start_year += 1
        connect.close()
        #将数据打包为字典，再返回
        Data = {'ID':log_ID,
                'Name':log_name,
                'Date':log_date,
                'PunchCard':log_punchcard,
                'Late':log_late,
                'LateTime':log_latetime,
                'Leave':log_leave,
                'LeaveTime':log_leavetime,
                'WorkTime':log_worktime,
                'WorkStatus':log_workstatus
                }
        return Data

    #模糊查询
    def Obscure_Search(self,start_date,end_date,emotion,content):
        #小月号数
        Small_Month = [4,6,9,11]
        #大月号数
        Big_Month = [1,3,5,7,8,10,12]
        #中月号数
        Middle_Month = [2]
        start_year = int(start_date[0])
        start_month = int(start_date[1])
        start_day = int(start_date[2])
        deadline = "{}-{}-{}".format(end_date[0],end_date[1],end_date[2])
        connect = sqlite3.connect(self.path)
        log_ID = []
        log_name = []
        log_date = []
        log_punchcard = []
        log_late = []
        log_latetime = []
        log_leave = []
        log_leavetime = []
        log_worktime = []
        log_workstatus = []
        while True:
            cursor = connect.cursor()
            table = "LogInfo_"+"{}".format(start_year)+"_"+"{}".format(start_month)
            if emotion == "None":
                if content == "迟到":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Late=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content))
                elif content == "早退":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Leave=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content))
                else:
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),))
            else:
                if content == "迟到":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Late=? AND WorkStatus=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),content,emotion))
                elif content == "早退":
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND Leave=? AND WorkStatus=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),contentn,emotion))
                else:
                    SQL = "SELECT ID,Name,Date_Info,PunchCard,Late,LateTime,Leave,LeaveTime,WorkTime,WorkStatus FROM {} WHERE Date_Info=? AND WorkStatus=?".format(table)
                    cursor.execute(SQL,("{}-{}-{}".format(start_year,start_month,start_day),emotion))
            all_data = cursor.fetchall()
            cursor.close()
            for row in all_data:
                log_ID.append(row[0])
                log_name.append(row[1])
                log_date.append(row[2])
                log_punchcard.append(row[3])
                log_late.append(row[4])
                log_latetime.append(row[5])
                log_leave.append(row[6])
                log_leavetime.append(row[7])
                log_worktime.append(row[8])
                log_workstatus.append(row[9])
            #到了检索终止日期就退出循环
            if "{}-{}-{}".format(start_year,start_month,start_day) == deadline:
                break 
            start_day += 1
            if start_month in Small_Month and start_day == 31:
                start_day = 1
                start_month += 1
            elif start_month in Big_Month and start_day == 32:
                start_day = 1
                start_month += 1
            elif start_month in Middle_Month and start_day == 29:
                #判断不是闰年
                if start_year % 4 == 0 and start_year % 100 == 0 and start_year % 400 != 0:
                    start_day = 1
                    start_month += 1
                #判断不是闰年
                elif start_year % 4 != 0:
                    start_day = 1
                    start_month += 1
                #判断是闰年
                else:
                    pass
            #闰年间二月的情况处理
            elif start_month in Middle_Month and start_day == 30:
                start_day = 1
                start_month += 1
            if start_month == 13:
                start_month = 1
                start_year += 1
        connect.close()
        #将数据打包为字典，再返回
        Data = {'ID':log_ID,
                'Name':log_name,
                'Date':log_date,
                'PunchCard':log_punchcard,
                'Late':log_late,
                'LateTime':log_latetime,
                'Leave':log_leave,
                'LeaveTime':log_leavetime,
                'WorkTime':log_worktime,
                'WorkStatus':log_workstatus
                }
        return Data    

#语音合成播报模块
class AISpeech:
    def __init__(self):
        self.ID = '16462683'
        self.KEY = 'biktygHgwFyTkZGML6E4a3Bb'
        self.SECRET_KEY = 'aICutwyohvZrnOD0vNhG1ZEeMLno9D0h'
    
    #发送一段字符串并获取该字符串的语音内容
    def SendText(self,text):
        Speech_Text = text
        Client = AipSpeech(self.ID, self.KEY, self.SECRET_KEY)
        Result = Client.synthesis(Speech_Text, 'zh', 1, {'spd':4 ,'vol':8 ,'per':0})
        if not isinstance(Result ,dict):
            with open('Speech.mp3','wb') as file:
                file.write(Result)

    #播放语音内容
    def Play(self):
        os.system('sox Speech.mp3 Speech.wav')
        read_file = wave.open('Speech.wav', 'rb')
        play = pyaudio.PyAudio()

        #异常信息反馈方法
        def callback(in_data, frame_count, time_info, status):
            data = read_file.readframes(frame_count)
            return (data, pyaudio.paContinue)
        #打开音频文件并获取其信息
        on_plya = play.open(format=play.get_format_from_width(read_file.getsampwidth()),
                               channels=read_file.getnchannels(),
                               rate=read_file.getframerate(),
                               output=True,
                               stream_callback=callback)
        #开始播放
        on_plya.start_stream()

        #循环检测是否播放完毕
        while on_plya.is_active():
            time.sleep(0.1)
        
        #播放完毕后执行以下代码
        on_plya.stop_stream()
        on_plya.close()
        read_file.close()
        play.terminate()

#自然语言分析模块
class AIParse:
    def __init__(self):
        self.ID = '16462683'
        self.KEY = 'biktygHgwFyTkZGML6E4a3Bb'
        self.SECRET_KEY = 'aICutwyohvZrnOD0vNhG1ZEeMLno9D0h'

    #DNN中文模型检测
    def CheckText(self,text):
        Client = AipNlp(self.ID, self.KEY, self.SECRET_KEY)
        Result = Client.dnnlm(text)
        Stander = float(Result.get('ppl'))
        if Stander > 200:
            return False
        else:
            return True

#服务器上传和下载
class ConnectServer:
    def __init__(self,ip,port):
        self.transport = paramiko.Transport((ip,port))
        self.username = "CloudDataBase"
        self.password = "yangshuyue1998"

    #连接至云服务器并创建一个SFTP实例
    def Connect(self):
        self.transport.connect(username=self.username,password=self.password)
        self.sftp = paramiko.SFTPClient.from_transport(self.transport)

    #将一个本地文件上传至云服务器
    def UpLoadFile(self,localpath,remotepath):
        self.sftp.put(localpath,remotepath,confirm=True)

    #从服务器获取一个文件并下载至本地
    def DownLoadFile(self,remotepath,localpath):
        self.sftp.get(remotepath,localpath)

    #关闭连接
    def CloseConnect(self):
        self.transport.close()

#上传、下载、导出、格式化进度条
class SimpleGauge(wx.Frame):
    def __init__(self,title,InfoText):
        wx.Frame.__init__(self,None,-1,title,size=(300,100),style=wx.CAPTION|wx.CLOSE_BOX)
        self.function = title
        self.InfoText = InfoText
        self.panel = wx.Panel(self,-1)
        self.panel.SetBackgroundColour('white')
        self.count = 0
        self.gauge = wx.Gauge(self.panel,-1,100,(20,20),(240,20))
        self.gauge.SetBezelFace(3)
        self.gauge.SetShadowWidth(3)
        self.Center(True)
        self.Bind(wx.EVT_IDLE, self.OnProcessDeal)       

    def OnProcessDeal(self,event): 
        time.sleep(0.05)
        self.count += 1
        self.gauge.SetValue(self.count)
        if self.count == 110:
            wx.MessageBox(self.function[0:2]+"完毕",caption="JoyYang提示")
            if self.function[0:2] == "上传":
                self.InfoText.AppendText(self.function[0:2]+"数据已保存至云端服务器\r\n")
            elif self.function[0:2] == "下载":
                self.InfoText.AppendText(self.function[0:2]+"数据已同步至本地\r\n")
            elif self.function[0:2] == "导出":
                self.InfoText.AppendText("文件已"+self.function[0:2]+"至本程序目录下Export文件夹\r\n")
            else:
                self.InfoText.AppendText("本地数据已经格式化完毕\r\n")
            self.gauge.Destroy()
            self.panel.Destroy()
            wx.Frame.Destroy(self)

#插入模拟打卡数据进度条
class ComplexGauge(wx.Frame):
    def __init__(self,date,employees,mintime,maxtime,leavingtime,database,infotext):
        wx.Frame.__init__(self,None,-1,"模拟数据生成中...",size=(300,100),style=wx.CAPTION|wx.CLOSE_BOX)
        self.count = 0
        self.date = date
        self.employees = employees
        self.MinTime = mintime
        self.MaxTime = maxtime
        self.LeavingTime = leavingtime
        self.database = database
        self.InfoText = infotext
        self.panel = wx.Panel(self,-1)
        self.panel.SetBackgroundColour('white')
        self.gauge = wx.Gauge(self.panel,-1,1500,(20,20),(240,20))
        self.gauge.SetBezelFace(3)
        self.gauge.SetShadowWidth(3)
        self.Center(True)
        if "月" in date and "年" in date:
            self.year = date[:date.find("年")]
            self.month = date[date.find("年")+1:date.find("月")]
            self.Bind(wx.EVT_IDLE, self.Month)
        else:
            self.year = date[:date.find("年")]
            self.Bind(wx.EVT_IDLE, self.Year)
    
    #自动生成某年某月的打卡数据
    def Month(self,event):
        #分离时分秒各个参数，转换为int类型，方便操作
        if self.MinTime[0] == "0":
            minhour = int(self.MinTime[1])
        else:
            minhour = int(self.MinTime[0:2])
        if self.MinTime[3] == "0":
            minute = int(self.MinTime[4])
        else:
            minute = int(self.MinTime[3:5])
        if self.MaxTime[0] == "0":
            maxhour = int(self.MaxTime[1])
        else:
            maxhour = int(self.MaxTime[0:2])      
        if self.LeavingTime[0] == "0":
            leave_hour = int(self.LeavingTime[1])
        else:
            leave_hour = int(self.LeavingTime[0:2])
        if self.LeavingTime[3] == "0":
            leave_minute = int(self.LeavingTime[4])
        else:
            leave_minute = int(self.LeavingTime[3:5])
        #创建一个工作情绪列表
        Emotion = ["高兴","舒坦","愉悦","闲适","平常","紧张","惆怅","消沉","沮丧","不快"]
        table = "LogInfo_"+self.year+"_"+self.month
        for i in range(31):
            for employee in self.employees:
                record = []
                ID = employee[0:11]
                Name = employee[11:]
                Date = self.year+"-"+self.month+"-"+"{}".format(i+1)
                #概率因子（设计为10%的迟到率）
                probability = random.choice(list(range(100)))
                #模拟判定为迟到
                if probability < 10:
                    if probability < 5:
                        carry_minute = int((minute + random.choice(list(range(10))) + probability * 10) / 60)
                        remian_minute = (minute + random.choice(list(range(10))) + probability * 10) % 60
                    else:
                        carry_minute = int((minute + random.choice(list(range(10))) + probability) / 60)
                        remian_minute = (minute + random.choice(list(range(10))) + probability) % 60                    
                    temp_hour = maxhour + carry_minute
                    temp_minute = remian_minute
                    temp_second = random.choice(list(range(60)))
                    #23:XX:XX的特殊情况
                    if temp_hour == 24:
                        temp_hour = 0
                    if len(str(temp_hour)) != 2:
                        temp_hour = "0{}".format(temp_hour)
                    else:
                        temp_hour = str(temp_hour)
                    if len(str(temp_minute)) != 2:
                        temp_minute = "0{}".format(temp_minute)
                    else:
                        temp_minute = str(temp_minute)
                    if len(str(temp_second)) != 2:
                        temp_second = "0{}".format(temp_second)
                    else:
                        temp_second = str(temp_second)
                    PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                    Late = "迟到"
                    LateTime = return_time_sustained(self.MaxTime,PunchCard)
                #模拟判定为未迟到
                else:
                    while True:
                        if maxhour - minhour <= 2:
                            temp_hour = random.choice([minhour,minhour+1,minhour+2])
                            temp_minute = random.choice(list(range(60)))
                            temp_second = random.choice(list(range(60)))
                            if len(str(temp_hour)) != 2:
                                temp_hour = "0{}".format(temp_hour)
                            else:
                                temp_hour = str(temp_hour)
                            if len(str(temp_minute)) != 2:
                                temp_minute = "0{}".format(temp_minute)
                            else:
                                temp_minute = str(temp_minute)
                            if len(str(temp_second)) != 2:
                                temp_second = "0{}".format(temp_second)
                            else:
                                temp_second = str(temp_second)
                            PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                            if self.MinTime <= PunchCard <= self.MaxTime:
                                break
                        #00:XX:XX - 23:XX:XX的特殊情况
                        elif maxhour - minhour == -23:
                            temp_hour = random.choice([minhour,0])
                            temp_minute = random.choice(list(range(60)))
                            temp_second = random.choice(list(range(60)))
                            if len(str(temp_hour)) != 2:
                                temp_hour = "0{}".format(temp_hour)
                            else:
                                temp_hour = str(temp_hour)
                            if len(str(temp_minute)) != 2:
                                temp_minute = "0{}".format(temp_minute)
                            else:
                                temp_minute = str(temp_minute)
                            if len(str(temp_second)) != 2:
                                temp_second = "0{}".format(temp_second)
                            else:
                                temp_second = str(temp_second)
                            PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                            if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "01:00:00"):
                                break
                        #00:XX:XX - 22:XX:XX的特殊情况
                        elif maxhour - minhour == -22 and minhour == 22:
                            temp_hour = random.choice([minhour,minhour+1,0])
                            temp_minute = random.choice(list(range(60)))
                            temp_second = random.choice(list(range(60)))
                            if len(str(temp_hour)) != 2:
                                temp_hour = "0{}".format(temp_hour)
                            else:
                                temp_hour = str(temp_hour)
                            if len(str(temp_minute)) != 2:
                                temp_minute = "0{}".format(temp_minute)
                            else:
                                temp_minute = str(temp_minute)
                            if len(str(temp_second)) != 2:
                                temp_second = "0{}".format(temp_second)
                            else:
                                temp_second = str(temp_second)
                            PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                            if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "01:00:00"):
                                break                       
                        #01:XX:XX - 23:XX:XX的特殊情况
                        elif maxhour - minhour == -22 and minhour == 23:
                            temp_hour = random.choice([minhour,0,1])
                            temp_minute = random.choice(list(range(60)))
                            temp_second = random.choice(list(range(60)))
                            if len(str(temp_hour)) != 2:
                                temp_hour = "0{}".format(temp_hour)
                            else:
                                temp_hour = str(temp_hour)
                            if len(str(temp_minute)) != 2:
                                temp_minute = "0{}".format(temp_minute)
                            else:
                                temp_minute = str(temp_minute)
                            if len(str(temp_second)) != 2:
                                temp_second = "0{}".format(temp_second)
                            else:
                                temp_second = str(temp_second)
                            PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                            if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "02:00:00"):
                                break
                    Late = "未迟到"
                    LateTime = "无迟到时间"
                #概率因子（设计为10%的早退率）
                probability = random.choice(list(range(100)))
                #模拟判定为早退
                if probability > 90:
                    probability = probability % 90
                    if probability >= 7:
                        temp_leave_hour = leave_hour - 1
                        #00:XX:XX的特殊情况
                        if temp_leave_hour == -1:
                            temp_leave_hour = 23
                        temp_leave_minute = leave_minute + random.choice(list(range(10)))
                        #处理分位数组大于等于60的情况
                        if temp_leave_minute >= 60:
                            temp_leave_minute = temp_leave_minute % 60
                            temp_leave_hour += 1
                        #23:XX:XX的特殊情况
                        if temp_leave_hour == 24:
                            temp_leave_hour = 0
                    else:
                        temp_leave_hour = leave_hour
                        temp_leave_minute = leave_minute - random.choice(list(range(10)))
                        #处理分位不够借时位再相减的情况
                        if temp_leave_minute < 0:
                            temp_leave_hour -= 1
                            temp_leave_minute += 60
                        #00:XX:XX的特殊情况
                        if temp_leave_hour == -1:
                            temp_leave_hour = 23
                    temp_leave_second = random.choice(list(range(60)))
                    if len(str(temp_leave_hour)) != 2:
                        temp_leave_hour = "0{}".format(temp_leave_hour)
                    else:
                        temp_leave_hour = str(temp_leave_hour)
                    if len(str(temp_leave_minute)) != 2:
                        temp_leave_minute = "0{}".format(temp_leave_minute)
                    else:
                        temp_leave_minute = str(temp_leave_minute)
                    if len(str(temp_leave_second)) != 2:
                        temp_leave_second = "0{}".format(temp_leave_second)
                    else:
                        temp_leave_second = str(temp_leave_second)
                    Leave = "早退"
                    LeaveTime = "{}:{}:{}".format(temp_leave_hour,temp_leave_minute,temp_leave_second)
                    WorkTime = return_time_sustained(PunchCard,LeaveTime)
                #模拟判定为未早退
                else:
                    if 10 <= probability <= 90:
                        temp_leave_hour = leave_hour
                        temp_leave_minute = leave_minute + random.choice(list(range(10)))
                        #处理分位数组大于等于60的情况
                        if temp_leave_minute >= 60:
                            temp_leave_minute = temp_leave_minute % 60
                            temp_leave_hour += 1
                        #23:XX:XX的特殊情况
                        if temp_leave_hour == 24:
                            temp_leave_hour = 0
                    else:
                        temp_leave_hour = leave_hour + 1
                        #23:XX:XX的特殊情况
                        if temp_leave_hour == 24:
                            temp_leave_hour = 0
                        temp_leave_minute = leave_minute - random.choice(list(range(10)))
                        #处理分位不够借时位再相减的情况
                        if temp_leave_minute < 0:
                            temp_leave_hour -= 1
                            temp_leave_minute += 60
                        #00:XX:XX的特殊情况
                        if temp_leave_hour == -1:
                            temp_leave_hour = 23
                    temp_leave_second = random.choice(list(range(60)))
                    if len(str(temp_leave_hour)) != 2:
                        temp_leave_hour = "0{}".format(temp_leave_hour)
                    else:
                        temp_leave_hour = str(temp_leave_hour)
                    if len(str(temp_leave_minute)) != 2:
                        temp_leave_minute = "0{}".format(temp_leave_minute)
                    else:
                        temp_leave_minute = str(temp_leave_minute)
                    if len(str(temp_leave_second)) != 2:
                        temp_leave_second = "0{}".format(temp_leave_second)
                    else:
                        temp_leave_second = str(temp_leave_second)
                    Leave = "未早退"
                    LeaveTime = "{}:{}:{}".format(temp_leave_hour,temp_leave_minute,temp_leave_second)
                    WorkTime = return_time_sustained(PunchCard,LeaveTime)
                WorkStatus = random.choice(Emotion)
                record.append(ID)
                record.append(Name)
                record.append(Date)
                record.append(PunchCard)
                record.append(Late)
                record.append(LateTime)
                record.append(Leave)
                record.append(LeaveTime)
                record.append(WorkTime)
                record.append(WorkStatus)
                self.database.Insert_Data(table,record)
                self.count += 1
                self.gauge.SetValue(self.count)
        wx.MessageBox("数据生成完毕",caption="JoyYang提示")
        self.InfoText.AppendText(self.date+"的打卡数据自动生成完毕\r\n请进行查询查看\r\n")
        self.gauge.Destroy()
        self.panel.Destroy()
        wx.Frame.Destroy(self)
    
    #自动生成某年的打卡数据
    def Year(self,event):
        #分离时分秒各个参数，转换为int类型，方便操作
        if self.MinTime[0] == "0":
            minhour = int(self.MinTime[1])
        else:
            minhour = int(self.MinTime[0:2])
        if self.MinTime[3] == "0":
            minute = int(self.MinTime[4])
        else:
            minute = int(self.MinTime[3:5])
        if self.MaxTime[0] == "0":
            maxhour = int(self.MaxTime[1])
        else:
            maxhour = int(self.MaxTime[0:2])      
        if self.LeavingTime[0] == "0":
            leave_hour = int(self.LeavingTime[1])
        else:
            leave_hour = int(self.LeavingTime[0:2])
        if self.LeavingTime[3] == "0":
            leave_minute = int(self.LeavingTime[4])
        else:
            leave_minute = int(self.LeavingTime[3:5])
        #1500为进度条上限值，现用1500来描述18000的数据总量，需每循环12次使self.count += 1
        num = 0
        #创建一个工作情绪列表
        Emotion = ["高兴","舒坦","愉悦","闲适","平常","紧张","惆怅","消沉","沮丧","不快"]
        for i in range(12):
            Table = "LogInfo_"+str(LocalTime.year)+"_"+"{}".format(i+1)
            for j in range(31):
                for employee in self.employees:
                    num += 1
                    record = []
                    ID = employee[0:11]
                    Name = employee[11:]
                    Date = self.year+"-"+"{}".format(i+1)+"-"+"{}".format(j+1)
                    #概率因子（设计为10%的迟到率）
                    probability = random.choice(list(range(100)))
                    #模拟判定为迟到
                    if probability < 10:
                        if probability < 5:
                            carry_minute = int((minute + random.choice(list(range(10))) + probability * 10) / 60)
                            remian_minute = (minute + random.choice(list(range(10))) + probability * 10) % 60
                        else:
                            carry_minute = int((minute + random.choice(list(range(10))) + probability) / 60)
                            remian_minute = (minute + random.choice(list(range(10))) + probability) % 60                    
                        temp_hour = maxhour + carry_minute
                        temp_minute = remian_minute
                        temp_second = random.choice(list(range(60)))
                        #23:XX:XX的特殊情况
                        if temp_hour == 24:
                            temp_hour = 0
                        if len(str(temp_hour)) != 2:
                            temp_hour = "0{}".format(temp_hour)
                        else:
                            temp_hour = str(temp_hour)
                        if len(str(temp_minute)) != 2:
                            temp_minute = "0{}".format(temp_minute)
                        else:
                            temp_minute = str(temp_minute)
                        if len(str(temp_second)) != 2:
                            temp_second = "0{}".format(temp_second)
                        else:
                            temp_second = str(temp_second)
                        PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                        Late = "迟到"
                        LateTime = return_time_sustained(self.MaxTime,PunchCard)
                    #模拟判定为未迟到
                    else:
                        while True:
                            if maxhour - minhour <= 2:
                                temp_hour = random.choice([minhour,minhour+1,minhour+2])
                                temp_minute = random.choice(list(range(60)))
                                temp_second = random.choice(list(range(60)))
                                if len(str(temp_hour)) != 2:
                                    temp_hour = "0{}".format(temp_hour)
                                else:
                                    temp_hour = str(temp_hour)
                                if len(str(temp_minute)) != 2:
                                    temp_minute = "0{}".format(temp_minute)
                                else:
                                    temp_minute = str(temp_minute)
                                if len(str(temp_second)) != 2:
                                    temp_second = "0{}".format(temp_second)
                                else:
                                    temp_second = str(temp_second)
                                PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                                if self.MinTime <= PunchCard <= self.MaxTime:
                                    break
                            #00:XX:XX - 23:XX:XX的特殊情况
                            elif maxhour - minhour == -23:
                                temp_hour = random.choice([minhour,0])
                                temp_minute = random.choice(list(range(60)))
                                temp_second = random.choice(list(range(60)))
                                if len(str(temp_hour)) != 2:
                                    temp_hour = "0{}".format(temp_hour)
                                else:
                                    temp_hour = str(temp_hour)
                                if len(str(temp_minute)) != 2:
                                    temp_minute = "0{}".format(temp_minute)
                                else:
                                    temp_minute = str(temp_minute)
                                if len(str(temp_second)) != 2:
                                    temp_second = "0{}".format(temp_second)
                                else:
                                    temp_second = str(temp_second)
                                PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                                if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "01:00:00"):
                                    break
                            #00:XX:XX - 22:XX:XX的特殊情况
                            elif maxhour - minhour == -22 and minhour == 22:
                                temp_hour = random.choice([minhour,minhour+1,0])
                                temp_minute = random.choice(list(range(60)))
                                temp_second = random.choice(list(range(60)))
                                if len(str(temp_hour)) != 2:
                                    temp_hour = "0{}".format(temp_hour)
                                else:
                                    temp_hour = str(temp_hour)
                                if len(str(temp_minute)) != 2:
                                    temp_minute = "0{}".format(temp_minute)
                                else:
                                    temp_minute = str(temp_minute)
                                if len(str(temp_second)) != 2:
                                    temp_second = "0{}".format(temp_second)
                                else:
                                    temp_second = str(temp_second)
                                PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                                if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "01:00:00"):
                                    break                       
                            #01:XX:XX - 23:XX:XX的特殊情况
                            elif maxhour - minhour == -22 and minhour == 23:
                                temp_hour = random.choice([minhour,0,1])
                                temp_minute = random.choice(list(range(60)))
                                temp_second = random.choice(list(range(60)))
                                if len(str(temp_hour)) != 2:
                                    temp_hour = "0{}".format(temp_hour)
                                else:
                                    temp_hour = str(temp_hour)
                                if len(str(temp_minute)) != 2:
                                    temp_minute = "0{}".format(temp_minute)
                                else:
                                    temp_minute = str(temp_minute)
                                if len(str(temp_second)) != 2:
                                    temp_second = "0{}".format(temp_second)
                                else:
                                    temp_second = str(temp_second)
                                PunchCard = "{}:{}:{}".format(temp_hour,temp_minute,temp_second)
                                if self.MinTime <= PunchCard or ("00:00:00" <= PunchCard < "02:00:00"):
                                    break
                        Late = "未迟到"
                        LateTime = "无迟到时间"
                    #概率因子（设计为10%的早退率）
                    probability = random.choice(list(range(100)))
                    #模拟判定为早退
                    if probability > 90:
                        probability = probability % 90
                        if probability >= 7:
                            temp_leave_hour = leave_hour - 1
                            #00:XX:XX的特殊情况
                            if temp_leave_hour == -1:
                                temp_leave_hour = 23
                            temp_leave_minute = leave_minute + random.choice(list(range(10)))
                            #处理分位数组大于等于60的情况
                            if temp_leave_minute >= 60:
                                temp_leave_minute = temp_leave_minute % 60
                                temp_leave_hour += 1
                            #23:XX:XX的特殊情况
                            if temp_leave_hour == 24:
                                temp_leave_hour = 0
                        else:
                            temp_leave_hour = leave_hour
                            temp_leave_minute = leave_minute - random.choice(list(range(10)))
                            #处理分位不够借时位再相减的情况
                            if temp_leave_minute < 0:
                                temp_leave_hour -= 1
                                temp_leave_minute += 60
                            #00:XX:XX的特殊情况
                            if temp_leave_hour == -1:
                                temp_leave_hour = 23
                        temp_leave_second = random.choice(list(range(60)))
                        if len(str(temp_leave_hour)) != 2:
                            temp_leave_hour = "0{}".format(temp_leave_hour)
                        else:
                            temp_leave_hour = str(temp_leave_hour)
                        if len(str(temp_leave_minute)) != 2:
                            temp_leave_minute = "0{}".format(temp_leave_minute)
                        else:
                            temp_leave_minute = str(temp_leave_minute)
                        if len(str(temp_leave_second)) != 2:
                            temp_leave_second = "0{}".format(temp_leave_second)
                        else:
                            temp_leave_second = str(temp_leave_second)
                        Leave = "早退"
                        LeaveTime = "{}:{}:{}".format(temp_leave_hour,temp_leave_minute,temp_leave_second)
                        WorkTime = return_time_sustained(PunchCard,LeaveTime)
                    #模拟判定为未早退
                    else:
                        if 10 <= probability <= 90:
                            temp_leave_hour = leave_hour
                            temp_leave_minute = leave_minute + random.choice(list(range(10)))
                            #处理分位数组大于等于60的情况
                            if temp_leave_minute >= 60:
                                temp_leave_minute = temp_leave_minute % 60
                                temp_leave_hour += 1
                            #23:XX:XX的特殊情况
                            if temp_leave_hour == 24:
                                temp_leave_hour = 0
                        else:
                            temp_leave_hour = leave_hour + 1
                            #23:XX:XX的特殊情况
                            if temp_leave_hour == 24:
                                temp_leave_hour = 0
                            temp_leave_minute = leave_minute - random.choice(list(range(10)))
                            #处理分位不够借时位再相减的情况
                            if temp_leave_minute < 0:
                                temp_leave_hour -= 1
                                temp_leave_minute += 60
                            #00:XX:XX的特殊情况
                            if temp_leave_hour == -1:
                                temp_leave_hour = 23
                        temp_leave_second = random.choice(list(range(60)))
                        if len(str(temp_leave_hour)) != 2:
                            temp_leave_hour = "0{}".format(temp_leave_hour)
                        else:
                            temp_leave_hour = str(temp_leave_hour)
                        if len(str(temp_leave_minute)) != 2:
                            temp_leave_minute = "0{}".format(temp_leave_minute)
                        else:
                            temp_leave_minute = str(temp_leave_minute)
                        if len(str(temp_leave_second)) != 2:
                            temp_leave_second = "0{}".format(temp_leave_second)
                        else:
                            temp_leave_second = str(temp_leave_second)
                        Leave = "未早退"
                        LeaveTime = "{}:{}:{}".format(temp_leave_hour,temp_leave_minute,temp_leave_second)
                        WorkTime = return_time_sustained(PunchCard,LeaveTime)
                    WorkStatus = random.choice(Emotion)
                    record.append(ID)
                    record.append(Name)
                    record.append(Date)
                    record.append(PunchCard)
                    record.append(Late)
                    record.append(LateTime)
                    record.append(Leave)
                    record.append(LeaveTime)
                    record.append(WorkTime)
                    record.append(WorkStatus)
                    self.database.Insert_Data(Table,record)
                    if num % 12 == 0:
                        num = num % 12
                        self.count += 1
                        self.gauge.SetValue(self.count)
        wx.MessageBox("数据生成完毕",caption="JoyYang提示")
        self.InfoText.AppendText(self.date+"的打卡数据自动生成完毕\r\n请进行查询查看\r\n")
        self.gauge.Destroy()
        self.panel.Destroy()
        wx.Frame.Destroy(self)