import time

import keyboard
import pyautogui

from config import PAUSE_SEC
from windows import wait_windows


def click_property(base_window):
    if wait_windows(base_window, time_check_second=99):
        keyboard.press_and_release('alt')
        keyboard.press_and_release('m')
        keyboard.press_and_release('b')


def click_num(num):
    pyautogui.click(1600, 299)
    keyboard.press_and_release('tab')
    des = num // 10

    keyboard.press_and_release(str(des))
    keyboard.press_and_release(str(des))

    for n in range(num % 10):
        keyboard.press_and_release('down')


def click_import(file, base_window):
    while True:
        if wait_windows(base_window, time_check_second=99):
            keyboard.press_and_release('alt')
            keyboard.press_and_release('m')
            keyboard.press_and_release('j')
            if wait_windows('Откр', time_check_second=2):
                break
    keyboard.write(file)
    time.sleep(PAUSE_SEC)
    keyboard.press_and_release('enter')
