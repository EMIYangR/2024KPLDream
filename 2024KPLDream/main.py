import GetDataFromGameClient
import SetJsonToExcel

if GetDataFromGameClient.getData() == 0:
    print("游戏端数据写入成功！")

if SetJsonToExcel.setJsonToExcel("data.json") == 0:
    print("赛事端数据写入成功！")
