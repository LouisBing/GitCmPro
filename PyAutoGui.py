import pyautogui, time, pyperclip

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

def paste(foo):
    pyperclip.copy(foo)
    pyautogui.hotkey('ctrl', 'v')

def locateCenterOnScreenWithTime(images,wait):
    # time.sleep(wait)
    sl = 0.3
    # v3.0：以次数为中心；循环20次（20次机会），i意义为次数减1，i和次数不统一；不再需要add；最终时间为+最大次数*0.5
    for i in range(0,20):
        for im, image in enumerate(images):
            if (pyautogui.locateCenterOnScreen(image,grayscale=True)):
                x, y = pyautogui.locateCenterOnScreen(image,grayscale=True)
                return x, y, im, wait + i * sl
                break
        time.sleep(sl)
    return -1, -1, -1, wait + 20 * sl

screenWidth, screenHeight = pyautogui.size()
pyautogui.moveTo(screenWidth / 2, screenHeight / 2)
print(screenWidth,screenHeight)

# currentMouseX, currentMouseY = pyautogui.position()
# pyautogui.moveTo(100, 500)
# pyautogui.click()
# #  鼠标向下移动10像素
# pyautogui.moveRel(None, 10)
# pyautogui.doubleClick()
# #  用缓动/渐变函数让鼠标2秒后移动到(500,1000)位置
# #  use tweening/easing function to move mouse over 2 seconds.
# pyautogui.moveTo(500, 1000, duration=2, tween=pyautogui.easeInOutQuad)
# #  在每次输入之间暂停0.25秒
# pyautogui.typewrite('Hello world!', interval=0.25)
# pyautogui.press('esc')
# pyautogui.keyDown('shift')
# pyautogui.press(['left', 'left', 'left', 'left', 'left', 'left'])
# pyautogui.keyUp('shift')
# pyautogui.hotkey('ctrl', 'c')

num = pyautogui.prompt(text='请复制处理意见，确保IE为活动窗口，确保有一个空白记事本在身边！', title='已阅！' , default='已阅。')
num = int(num)

# 待办第一条的位置坐标
tempPn = 6
pointOne = (95, 268+32*tempPn)

tTodo = 0.5
tN = 0.5
tUp = 0.5
tUpJt = 0.5
READ = '已阅。'

for i in range(num):
    # 检测是否返回待办页面，如没有直接终止。
    # im = pyautogui.screenshot()
    # if im.getpixel((300,180)) != (241,100,78):
    #     print("Not Back",300,180)
    #     break
    pointOne = (95, 268 + 32 * tempPn)
    # 政企点击的坐标序列：处理，已阅，结束，提交，确认
    pointDeals = ((181, 105), (604, 385), (789, 333), (714, 574), (670, 448))

    x, y, w, tTodo = locateCenterOnScreenWithTime(['todolist.png'], tTodo)
    pyautogui.press('f5')
    x, y, w, tTodo = locateCenterOnScreenWithTime(['todolist.png'], tTodo)
    if x==-1:
        pyautogui.alert(text='No Back', title='Error', button='OK')
        print("Not Back", 300, 180)
        break
    print('tTodo=', tTodo)
    pyautogui.click(pointOne)
    # pyautogui.click(pointOne,duration=1)

    # 下面是点击前延迟4.5秒，并不是说点击后延迟4.5秒
    # pyautogui.click(pointOne, duration=4.5)
    # 弹出正文时间3.5->4-5
    pyautogui.moveTo(y=0)
    x, y, which, tN = locateCenterOnScreenWithTime(['next-zq.png','next-zq2.png', 'next-jt.png'], tN)
    print('tN=',tN)
    # 政企处理方法
    if (which<2):
        pyautogui.click(x, y)
        # 弹出提交窗口时间优化：2.5->2,4-20
        x, y, w, tUp = locateCenterOnScreenWithTime(['up-zq.png'], tUp)
        print('tUp=', tUp)
        pyautogui.moveTo(pointDeals[1])
        # 点击文本框并粘贴
        pyautogui.click()
        paste(READ)
        # pyautogui.hotkey('ctrl', 'v')
        # 点击结束办理
        pyautogui.moveTo(pointDeals[2], duration=0.5)
        pyautogui.click()
        # 点击提交
        pyautogui.click(pointDeals[3])
        # 弹出确认窗口并确认
        pyautogui.moveTo(pointDeals[4], duration=1.5)
        pyautogui.click()
        # x, y, w, tOk = locateCenterOnScreenWithTime(['ok-zq.png'], 0)
        # pyautogui.click(x, y)

        # # 临时5.5-10
        # pyautogui.moveTo(x=screenWidth, duration=20)
    # 集团处理方法
    elif(which==2):
        tempPn +=1
        pyautogui.hotkey('ctrl', 'w')
        continue
        # 集团点击的坐标序列：处理，已阅，结束，提交，确认
        pointDeals = ((181, 105), (604, 385), (930, 228), (657, 621), (653, 260))
        pyautogui.click(x,y)
        # 弹出提交窗口时间4.5会出问题，变为5
        x, y, w, tUpJt = locateCenterOnScreenWithTime(['up-jt.png'], tUpJt)
        print('tUpJt=',tUpJt)
        pyautogui.moveTo(pointDeals[1])
        # 点击文本框并粘贴
        pyautogui.click()
        paste(READ)
        # pyautogui.hotkey('ctrl', 'v')
        # 点击结束办理
        pyautogui.moveTo(pointDeals[2], duration=1)
        pyautogui.click()
        # 点击提交
        pyautogui.click(pointDeals[3])
        # 弹出确认窗口并确认2有问题，变为2.5
        pyautogui.moveTo(pointDeals[4], duration=2.5)
        pyautogui.click()
        # 确认的调度与第一行待办在一个高度，防止鼠标停留改变第一行待办的颜色
        pyautogui.moveTo(y=0)
    else:
        pyautogui.alert(text='No Next', title='Error', button='OK')
        print('no next')
        break

    timeList = 'tTodo=%.2f,tN=%.2f,tUp=%.2f,tUpJt=%.2f' % (tTodo, tN, tUp, tUpJt)
    print(timeList)
    writeFile = 'PyAutoGui.txt'
    TxtOperator.writeList2Txt(writeFile, [timeList], 'a')

