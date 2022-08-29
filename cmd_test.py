# -*- coding: utf-8 -*-
import os
cmd = os.popen("c:\configureMosaic.exe listconfigcmd").read().split(" ")
rows = cmd[2].split("=")[-1]

firstgrid = cmd[2:5]
firstgrid.insert(0, 'firstgrid')
print(firstgrid)
try:
    nextgridIndex = cmd.index('nextgrid')
    print(cmd[nextgridIndex:nextgridIndex + 4])
except:
    pass
print(cmd)


