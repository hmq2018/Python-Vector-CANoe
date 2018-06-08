# Python-Vector-CANoe

Control Vector CANoe API by Python

Install:

pip install Python-CANoe



Usage:

app = CANoe.CANoe() #定义CANoe为app

app.open_simulation("test.cfg") #导入某个CANoe congif

app.start_Measurement() #启动CANoe

var_from_namespace = app.get_all_SysVar("mfl") #获取namespace下的所有系统变量

print(app.get_SysVar("mfl","vol_plus")) #获取系统变量的值

app.set_SysVar("mfl","vol_plus",1)  #写入系统变量的值

app.stop_Measurement() #停止CANoe
