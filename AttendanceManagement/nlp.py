from aip import AipNlp
class AIParse:
    def __init__(self):
        self.ID = '16462683'
        self.KEY = 'biktygHgwFyTkZGML6E4a3Bb'
        self.SECRET_KEY = 'aICutwyohvZrnOD0vNhG1ZEeMLno9D0h'

    #DNN中文模型检测
    def CheckText(self,text):
        Client = AipNlp(self.ID, self.KEY, self.SECRET_KEY)
        Result = Client.dnnlm(text)
        return Result

text = "查询杨某在2020年4月1日至2020年4月30日的迟到记录"
dnn = AIParse()
result = dnn.CheckText(text)
print(result)