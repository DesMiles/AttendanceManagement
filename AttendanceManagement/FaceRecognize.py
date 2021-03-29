import os
import csv
import datetime
import wx
import json
import _thread
import dlib             #人脸识别库dlib
import numpy as np      #数据处理库numpy
import cv2              #图像处理库OpenCv
import pandas as pd     #数据处理库pandas
import ApplicationTool

camera = 'http://192.168.3.23:8080/?action=stream?dummy=param.mjpg'

#加载人脸识别分类器
detector = dlib.get_frontal_face_detector()
predictor = dlib.shape_predictor("Classifiers/shape_predictor_face_landmarks.dat")
classifiers = dlib.face_recognition_model_v1("Classifiers/dlib_face_recognition_model.dat")

screen = "Picture/screen.png"
holder = "Picture/holder.png"
success = "Picture/success.png"
fail = "Picture/fail.png"
repeat = "Picture/repeat.png"
attendance = "Picture/attendance.png"

#数据库路径
path_dataBase = "./Data/loglist.db"
#人脸库路径
path_feature_all = "./Data/feature_all.csv"
#存放设置打卡时间的数据路径
path_setting_time = "./Data/setting_time.json"

#存放设置打卡时间的数据（如果不存在则创建）
if not os.path.exists(path_setting_time):
    SettingTime = {
        'MinTime':"07:00:00",
        'MaxTime':"09:00:00",
        'LeavingTime':"17:00:00"
        }
    json_object = json.dumps(SettingTime)
    with open(path_setting_time, 'w') as jsonfile:
             jsonfile.write(json_object)
             jsonfile.close()

#处理存放所有人脸特征的数据表格（如果不存在则创建）
if not os.path.exists(path_feature_all):
    with open(path_feature_all, 'w', newline='') as csvfile:
             csv_write = csv.writer(csvfile)
             csv_head = []
             for i in range(128):
                 csv_head.append(0)
             csv_head.append("ID")
             csv_head.append("Name")            
             csv_write.writerow(csv_head)

#用来存放所有录入人脸特征的数组
features_known_arr = []
csv_rd = pd.read_csv(path_feature_all,header=None,encoding='gbk')
for i in range(csv_rd.shape[0]):
    features_someone_arr = []
    for j in range(0, len(csv_rd.loc[i,:])):
        features_someone_arr.append(csv_rd.loc[i,:][j])
    features_known_arr.append(features_someone_arr)

#计算并返回两个人脸特征向量间的欧式距离（我论文中有相关知识点的介绍，看人脸识别算法那一节）
def return_euclidean_distance(feature_1,feature_2):
    feature_1 = np.array(feature_1)
    feature_2 = np.array(feature_2)
    dist = np.sqrt(np.sum(np.square(feature_1 - feature_2)))
    print("欧式距离: ", dist)
    if dist > 0.4:
        return "diff"
    else:
        return "same"

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
    sust = "{}小时-{}分钟".format(hour,minute)
    return sust

#打卡界面
class PunchcardUI(wx.Frame):
    def __init__(self,rut,InfoText,MinTime,MaxTime,LeavingTime,flag):
        wx.Frame.__init__(self,parent=rut,title="打卡界面",size=(840,620),style=wx.CAPTION|wx.CLOSE_BOX|wx.STAY_ON_TOP)
        wx.Frame.SetMinSize(self,size=(840,630))
        wx.Frame.SetMaxSize(self,size=(840,630))
        self.SetBackgroundColour('white')
        self.Center()       

        #定义各种参数
        self.database = ApplicationTool.DataBase(path_dataBase) #连接数据库
        self.InfoText = InfoText                                #文本控制台
        self.count = 0                                          #定时器
        self.MinTime = MinTime                                  #最早打卡的时间
        self.MaxTime = MaxTime                                  #最晚打卡的时间
        self.LeavingTime = LeavingTime                          #下班打卡的时间
        self.flag = flag                                        #两种打卡方式选择标志

        #文本信息栏输出当前人脸库人脸数量
        self.InfoText.AppendText("数据库人脸数:{}\r\n".format(len(features_known_arr) - 1))

        #图片像素比调整
        self.image_screen = wx.Image(screen, wx.BITMAP_TYPE_ANY).Scale(820,480)
        self.image_holder = wx.Image(holder, wx.BITMAP_TYPE_ANY).Scale(500,100)
        self.image_success = wx.Image(success, wx.BITMAP_TYPE_ANY).Scale(820,480)  
        self.image_fail = wx.Image(fail, wx.BITMAP_TYPE_ANY).Scale(820,480)
        self.image_repeat = wx.Image(repeat, wx.BITMAP_TYPE_ANY).Scale(820,480)    
        self.image_attendance = wx.Image(attendance,wx.BITMAP_TYPE_PNG).ConvertToBitmap()

        #初始化按键以及绑定事件
        self.StartButton = wx.BitmapButton(self,wx.ID_ANY,self.image_attendance,pos=(30,490),size=(100,100)) 
        self.StartButton.SetBackgroundColour('white')
        self.Bind(wx.EVT_BUTTON,self.OnButtonClicked,self.StartButton)

        #设置两个区域显示图片
        self.screen = wx.StaticBitmap(parent=self, pos=(0,0), bitmap=wx.Bitmap(self.image_screen))
        self.holder = wx.StaticBitmap(parent=self, pos=(160,480), bitmap=wx.Bitmap(self.image_holder))  
        
    #打开摄像头创建子线程
    def OnButtonClicked(self,event):       
        if self.flag == "punchcard":
            _thread.start_new_thread(self._open_cap_for_punchcard, (event,))
        elif self.flag == "leave":
            _thread.start_new_thread(self._open_cap_for_leave, (event,))

    def _open_cap_for_punchcard(self,event):
        #创建一个摄像头对象
        self.capture = cv2.VideoCapture(0)#参数0是调用电脑上默认摄像头
        #判断摄像头是否已经打开并循环检测
        while self.capture.isOpened():
            #定时器计时开始
            self.count += 1
            #cap.read()方法会返回两个值：
            #一个布尔值true/false，用来判断读取视频是否成功
            #另一个是图像对象，即图像的三维矩阵
            success, frame = self.capture.read()
            #摄像头异常时，直接跳出
            if not success:
                capture.release()
                cv2.destroyAllWindows()
                for i in list(range(0,4,1)):
                    cv2.waitKey(1)
                self.InfoText.AppendText("摄像头异常，请重试\r\n")
                break
            #人脸数
            face_num = detector(frame, 1)

            #人脸姓名标签初始化
            Name = ""

            #判断检测到了人脸
            if len(face_num) != 0:               
                #获取当前捕获到的图像的人脸的特征
                features_cap = ''
                shape = predictor(frame, face_num[0])
                features_cap = classifiers.compute_face_descriptor(frame, shape)

                #对应当前帧出现的人脸，遍历所有存储的人脸特征
                for i in range(len(features_known_arr)):
                    #将当前人脸与存储的所有人脸数据进行比对
                    compare = return_euclidean_distance(features_cap, features_known_arr[i][0:-2])
                    #找到了相似脸，即成功识别人脸
                    if compare == "same":
                        record = []
                        ID = features_known_arr[i][-2]
                        Name = features_known_arr[i][-1]                       
                        record.append(ID)
                        record.append(Name)
                        LocalTime = datetime.datetime.now()
                        Date = str(LocalTime.year)+"-"+str(LocalTime.month)+"-"+str(LocalTime.day)
                        time_hour = str(LocalTime.hour)
                        time_minute = str(LocalTime.minute)
                        time_second = str(LocalTime.second)
                        if len(time_hour) == 1:
                            time_hour = "0"+time_hour
                        if len(time_minute) == 1:
                            time_minute = "0"+time_minute
                        if len(time_second) == 1:
                            time_second = "0"+time_second
                        Time = time_hour+":"+time_minute+":"+time_second
                        record.append(Date)
                        record.append(Time)
                        if self.MinTime <= Time <= self.MaxTime:
                            puncard_info = self.database.FindPunchCardInfo(str(ID),Date)
                            #检测是否重复签到
                            if puncard_info == True:
                                self.screen.SetBitmap(wx.Bitmap(self.image_repeat))
                                self.InfoText.AppendText(Name+"今天已经进行过上班打卡\r\n")
                                self.InfoText.AppendText("已重复打卡警告\r\n")
                                wx.MessageBox("您今天已经进行过上班打卡", caption="JoyYang提示",parent=self)
                                self.capture.release()
                                cv2.destroyAllWindows()
                                for i in list(range(0,4,1)):
                                    cv2.waitKey(1)
                                _thread.exit()

                            self.InfoText.AppendText(Name+"上班打卡成功，且未迟到\r\n签到时间:"+Date+" "+Time+"\r\n")
                            #填入出勤情况
                            record.append("未迟到")
                            #填入迟到时长
                            record.append("无迟到时长")
                            #用Unknown填入剩下的字段
                            for i in range(4):
                                record.append("Unknown")

                        elif Time < self.MinTime:
                            self.screen.SetBitmap(wx.Bitmap(self.image_fail))
                            wx.MessageBox("还未到打卡时间进行打卡，打卡失败\n当前时间为:"+Time+"\n最早打卡时间为:"+self.MinTime, caption="JoyYang提示",parent=self)
                            self.InfoText.AppendText(Name+"还未到打卡时间进行打卡，打卡失败\r\n")
                            self.capture.release()
                            cv2.destroyAllWindows()
                            for i in list(range(0,4,1)):
                                cv2.waitKey(1)
                            _thread.exit()

                        elif Time >= self.LeavingTime:
                            puncard_info = self.database.FindPunchCardInfo(str(ID),Date)
                            if puncard_info == True:
                                self.screen.SetBitmap(wx.Bitmap(self.image_repeat))
                                wx.MessageBox("您今天已经进行过上班打卡\n现在已经是下班时间，可以进行下班打卡", caption="JoyYang提示",parent=self)
                                self.InfoText.AppendText(Name+"今天已经进行过上班打卡\r\n")
                                self.InfoText.AppendText("已重复打卡警告\r\n")
                                self.capture.release()
                                cv2.destroyAllWindows()
                                for i in list(range(0,4,1)):
                                    cv2.waitKey(1)
                                _thread.exit()
                            else:
                                self.screen.SetBitmap(wx.Bitmap(self.image_fail))
                                wx.MessageBox("您今天已经旷工，打卡失败，请明天及时向老板说明情况", caption="JoyYang提示",parent=self)
                                #填入出勤情况
                                record.append("旷工")
                                #用Unknown填入剩下的字段
                                for i in range(5):
                                    record.append("Unknown")
                        else:
                            #检测是否重复签到
                            puncard_info = self.database.FindPunchCardInfo(str(ID),Date)
                            if puncard_info == True:
                                self.screen.SetBitmap(wx.Bitmap(self.image_repeat))
                                self.capture.release()
                                cv2.destroyAllWindows()
                                for i in list(range(0,4,1)):
                                    cv2.waitKey(1)
                                _thread.exit()

                            self.InfoText.AppendText(Name+"上班打卡成功，但迟到了\r\n打卡时间:"+Date+" "+Time+"\r\n")
                            #填入出勤情况
                            record.append("迟到")
                            #填入迟到时长
                            record.append(return_time_sustained(self.MaxTime,Time))
                            #用Unknown填入剩下的字段
                            for i in range(4):
                                record.append("Unknown")

                        #将员工工号、员工姓名、打卡日期、上班打卡时间、出勤情况、迟到时长等信息写入数据库                        
                        self.database.Insert(record)
                        self.InfoText.AppendText("打卡信息已经写入数据库\r\n")  
                        self.screen.SetBitmap(wx.Bitmap(self.image_success)) 
                        self.capture.release()
                        cv2.destroyAllWindows()
                        for i in list(range(0,4,1)):
                            cv2.waitKey(1)
                        _thread.exit()
            
            self.screen.SetBitmap(wx.Bitmap("Picture/Wait/"+str(self.count%6)+".png"))

            #定时器达到定时上限
            if self.count == 18:
                self.count = 0
                self.screen.SetBitmap(wx.Bitmap(self.image_fail))
                self.capture.release()
                cv2.destroyAllWindows()
                for i in list(range(0,4,1)):
                    cv2.waitKey(1)
                _thread.exit()    

    def _open_cap_for_leave(self,event):
            #创建一个摄像头对象
            self.capture = cv2.VideoCapture(0)#参数0是调用电脑上默认摄像头
            #判断摄像头是否已经打开并循环检测
            while self.capture.isOpened():
                #定时器计时开始
                self.count += 1
                #cap.read()方法会返回两个值：
                #一个布尔值true/false，用来判断读取视频是否成功
                #另一个是图像对象，即图像的三维矩阵
                success, frame = self.capture.read()
                #摄像头异常时，直接跳出
                if not success:
                    capture.release()
                    cv2.destroyAllWindows()
                    for i in list(range(0,4,1)):
                        cv2.waitKey(1)
                    self.InfoText.AppendText("摄像头异常，请重试\r\n")
                    break
                #人脸数
                face_num = detector(frame, 1)

                #人脸姓名标签初始化
                Name = ""

                #判断检测到了人脸
                if len(face_num) != 0:               
                    #获取当前捕获到的图像的人脸的特征
                    features_cap = ''
                    shape = predictor(frame, face_num[0])
                    features_cap = classifiers.compute_face_descriptor(frame, shape)

                    #对应当前帧出现的人脸，遍历所有存储的人脸特征
                    for i in range(len(features_known_arr)):
                        #将当前人脸与存储的所有人脸数据进行比对
                        compare = return_euclidean_distance(features_cap, features_known_arr[i][0:-2])
                        #找到了相似脸，即成功识别人脸
                        if compare == "same":
                            update_info = []
                            ID = features_known_arr[i][-2]
                            Name = features_known_arr[i][-1] 
                            LocalTime = datetime.datetime.now()
                            Date = str(LocalTime.year)+"-"+str(LocalTime.month)+"-"+str(LocalTime.day)
                            time_hour = str(LocalTime.hour)
                            time_minute = str(LocalTime.minute)
                            time_second = str(LocalTime.second)
                            if len(time_hour) == 1:
                                time_hour = "0"+time_hour
                            if len(time_minute) == 1:
                                time_minute = "0"+time_minute
                            if len(time_second) == 1:
                                time_second = "0"+time_second
                            Time = time_hour+":"+time_minute+":"+time_second                       
                            puncard_info = self.database.FindPunchCardInfo(str(ID),Date)
                            if puncard_info == True and Time >= self.LeavingTime:                    
                                punchcard_time = self.database.FindPunchCardTime(str(ID),Date)
                                leave_info = self.database.FindLeaveInfo(str(ID),Date)
                                #检测是否重复签到
                                if leave_info == True:
                                    self.screen.SetBitmap(wx.Bitmap(self.image_repeat))
                                    wx.MessageBox("您今天已经进行过下班打卡", caption="JoyYang提示",parent=self)
                                    self.InfoText.AppendText(Name+"今天已经进行过下班打卡\r\n")
                                    self.capture.release()
                                    cv2.destroyAllWindows()
                                    for i in list(range(0,4,1)):
                                        cv2.waitKey(1)
                                    _thread.exit()

                                self.InfoText.AppendText(Name+"下班打卡成功\r\n打卡时间:"+Date+" "+Time+"\r\n")
                                #更新早退情况
                                update_info.append("未早退")
                                #更新下班打卡时间
                                update_info.append(Time)
                                #更新工作时长
                                update_info.append(return_time_sustained(punchcard_time,Time))
                            else:
                                if puncard_info == False:
                                    wx.MessageBox("下班打卡失败，请先确认是否进行上班打卡",caption="JoyYang提示",parent=self)
                                    self.screen.SetBitmap(wx.Bitmap(self.image_fail))
                                    self.InfoText.AppendText("下班打卡异常\r\n")
                                    self.capture.release()
                                    cv2.destroyAllWindows()
                                    for i in list(range(0,4,1)):
                                        cv2.waitKey(1)
                                    _thread.exit()
                                else:
                                    punchcard_time = self.database.FindPunchCardTime(str(ID),Date)
                                    leave_info = self.database.FindLeaveInfo(str(ID),Date)
                                    prompt = "还未到下班时间，下班时间为"+self.LeavingTime+",现在进行下班打卡的话会被记录为早退，是否继续？"
                                    dig = wx.MessageDialog(self,prompt,"JoyYang提示",wx.YES_NO|wx.ICON_QUESTION)

                                    if leave_info == True:
                                        self.screen.SetBitmap(wx.Bitmap(self.image_repeat))
                                        wx.MessageBox("您今天已经进行过下班打卡", caption="JoyYang提示",parent=self)
                                        self.InfoText.AppendText(Name+"今天已经进行过下班打卡\r\n")
                                        self.capture.release()
                                        cv2.destroyAllWindows()
                                        for i in list(range(0,4,1)):
                                            cv2.waitKey(1)
                                        _thread.exit()

                                    if dig.ShowModal() == wx.ID_YES:
                                        self.InfoText.AppendText(Name+"下班打卡成功\r\n打卡时间:"+Date+" "+Time+"\r\n")
                                        #更新早退情况
                                        update_info.append("早退")
                                        #更新下班打卡时间
                                        update_info.append(Time)
                                        #更新工作时长
                                        update_info.append(return_time_sustained(punchcard_time,Time))
                                    else:
                                        self.screen.SetBitmap(wx.Bitmap(self.image_screen))
                                        self.InfoText.AppendText("下班打卡已取消\r\n")
                                        self.capture.release()
                                        cv2.destroyAllWindows()
                                        for i in list(range(0,4,1)):
                                            cv2.waitKey(1)
                                        _thread.exit()

                            #将更新信息写入数据库
                            self.database.ModifyLeaveInfo(str(ID),Date,update_info)
                            self.InfoText.AppendText("更新信息已经写入数据库\r\n")
                            self.screen.SetBitmap(wx.Bitmap(self.image_success))                      

                            self.capture.release()
                            cv2.destroyAllWindows()
                            for i in list(range(0,4,1)):
                                cv2.waitKey(1)
                            _thread.exit()
            
                self.screen.SetBitmap(wx.Bitmap("Picture/Wait/"+str(self.count%6)+".png"))

                #定时器达到定时上限
                if self.count == 18:
                    self.count = 0
                    self.screen.SetBitmap(wx.Bitmap(self.image_fail))
                    self.capture.release()
                    cv2.destroyAllWindows()
                    for i in list(range(0,4,1)):
                        cv2.waitKey(1)
                    _thread.exit()