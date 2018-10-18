

import time
import os
from subprocess import call

counter_global = 0
counter_daily = 0
today = time.localtime()[2]

while True:

    if today != time.localtime()[2]:

        today = time.localtime()[2]
        counter_daily = 0
        
    t_i = time.localtime()
    print "Restart number (ALL): ", counter_global
    print "Restart number (today): ", counter_daily
    print "Restarting at (Y, M, D): ", t_i[0], t_i[1], t_i[2]
    print "Restarting at (h, m, s): ", t_i[3], t_i[4], t_i[5]
    os.system("taskkill /F /IM EXCEL.EXE")
    call(["python", "PyCor.pyc"])
    print "-------------------------"

    counter_global += 1
    counter_daily += 1

    time.sleep(15)
