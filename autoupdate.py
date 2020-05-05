from pywinauto import application
from pywinauto import timings
from pywinauto import findwindows
import time
import os
import json

with open('C:/Users/user/Projects/kw.crd') as f:
    data = json.load(f)

app = application.Application()
app.start("C:/KiwoomFlash3/Bin/NKMiniStarter.exe")

title = "번개3 Login"
dlg = timings.WaitUntilPasses(20, 0.5, lambda: app.window_(title=title))

pass_ctrl = dlg.Edit2
pass_ctrl.SetFocus()
pass_ctrl.TypeKeys(data['kw'])

cert_ctrl = dlg.Edit3
cert_ctrl.SetFocus()
cert_ctrl.TypeKeys('')

btn_ctrl = dlg.Button0
btn_ctrl.Click()

time.sleep(1)

dlg2 = timings.WaitUntilPasses(20, 0.5, lambda: app.window_(title = '번개3'))
btn_ctrl2 = dlg2.Button1
btn_ctrl2.Click()

time.sleep(20)
os.system("taskkill /im NKmini.exe")

time.sleep(1)

title2 = findwindows.find_windows(title='', class_name='#32770', control_id=0)[0]
dlg2 = timings.WaitUntilPasses(20, 0.5, lambda: app.window_(handle=title2))
btn_ctrl2 = dlg2.Button0
btn_ctrl2.Click()

