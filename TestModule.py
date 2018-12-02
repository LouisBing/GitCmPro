#-*-coding:utf-8-*-
# ------------------------------------------------------------------------------------------
import pyautogui,pyperclip
pyautogui.PAUSE = 0.1
while(True):
    num = pyautogui.prompt(text='请复制处理意见，确保IE为活动窗口，确保有一个空白记事本在身边！', title='已阅！' , default='20')
    print(num)
    num = int(num)
    pyperclip.copy(num)
    x, y = pyautogui.position()
    pyautogui.click((x,y))
    pyautogui.hotkey('ctrl', 'v')
    for i in range(50):
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'v')
        20
    # print(num)