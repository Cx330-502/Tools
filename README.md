# Tools
### Schedule

这是一个自动生成排班表的脚本~

#### 使用说明

> 读取的excel文件文件名必须是sche.xlsx ！！！

```shell
pip install requirements.txt
python schedule.txt
```

根据提示选择目录后等待即可。

输入文件格式应当如下：![pic1](./images/pic1.png)

工作表数量应当有两张，格式相同。

#### 输出说明

由于整个软件运行时间较长，因此每一分钟会输出一个中间文件，用户可以选择仅使用中间文件。

每隔五分钟用户可选择是否中断当前进程，直接进入下一进程。

#### 打包说明

```shell
pip install pyinstaller
pyinstaller -F --onefile --name=sche schedule.py --hidden-import=openpyxl.cell._writer
```

`exe` 文件将生成在 `dist` 文件夹下 

#### 算法说明

由于一开始不知道 pyinstaller 怎么打包多文件，所以把所有东西都粘到一个文件里了，然后后来再改也都是直接在这个基础上修改 . . . . . .

一开始尝试过穷举，发现运行时间实在太长了，完全不可取，然后更换了算法。

目前采取的是通过四轮调度来获取尽可能更好的排班。

第一轮调度 schedule1 先遍历所有的时间段，若某一时间段内可用人数较少则优先选中。

第二轮调度 schedule2 再次遍历所有时间段，在每人的时长不太长的情况下尽可能多的在每个时间段安排人员。

第三轮调度 schedule3 遍历所有人员，将调度时间较少的人员安排到可参与的时间段。

第四轮调度 schedule4 再次遍历所有人员，将调度时间较长的人员的工作时间尽可能缩减。

第五轮调度 schedule5 重复了第四轮的过程。

实测第二、三轮调度耗时较长orz

如果有更好的算法可以在 issue 里提出😭😭😭
