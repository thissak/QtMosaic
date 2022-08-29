#스레드 종료 대기

import threading
import time
import random
import os

def DoItThread(str):
    cnt = 0
    while(cnt<10):
        time.sleep(random.randint(0,100)/300.0)
        print(str,cnt)
        cnt+=1
    print("=== ",str,"스레드 종료 ===")

def enableMosaic():
    xmlString = os.popen(
        "C:\\configureMosaic.exe test rows=1 cols=1 res=2560,1600,165.000 gridPos=0,0 out=0,0 rotate=0  nextgrid rows=1 cols=1 res=3840,2160,60.000 gridPos=2560,0 out=0,1 rotate=0").read()
    print(xmlString, "스레드 종료 ===")

# th_a = threading.Thread(target = DoItThread, args=("홍길동",))
th_b = threading.Thread(target = enableMosaic)

print("=== 스레드 가동 ===")
# th_a.start()
th_b.start()

# th_a.join()
th_b.join()
print("테스트 종료")