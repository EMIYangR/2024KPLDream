# 2024KPLDream

本项目基于Python和JavaScript，运行main.py即可将2024年KPL梦之队的游戏端投票数据加载并写入到示例Excel中，你可以自定义值来自行选择你要写入的数据。

本次更新基于

[Fillder]: https://www.telerik.com/download/fiddler	"Fillder"

抓包获取Json格式数据，你可以自行保存它到“data.json”文件中（需要自行创建），然后运行SetJsonToExcel.py加载到Excel中，你也可以配置好后直接运行main.py一键加载获取的赛事端数据并拉取游戏端数据进行写入。

## 默认值：

### 选手

|  值   |  对应值  |
| :---: | :------: |
| nick  |   昵称   |
| road  |   分路   |
| vote  |   票数   |
| tname | 所属队伍 |
|  mid  |    ID    |

### 教练

|    值    |    对应值    |
| :------: | :----------: |
|   name   | 昵称（姓名） |
| votevote |     票数     |
|    id    |      ID      |

## 可选值

### 选手

|  值   |  对应值  |
| :---: | :------: |
| tname | 所属队伍 |
| nick  |   昵称   |
| name  |   姓名   |
| road  |   分路   |
| vote  |   票数   |

### 教练

|    值    |    对应值    |
| :------: | :----------: |
|   name   | 昵称（姓名） |
| votevote |     票数     |

