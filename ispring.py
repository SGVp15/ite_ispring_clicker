import time

import keyboard
import pyautogui

from config import PAUSE_SEC, WINDOW_NAME_OPEN, WINDOW_NAME_PROPERTY
from windows import wait_windows


def click_property(base_window):
    while True:
        if wait_windows(base_window, time_check_second=99):
            keyboard.press_and_release('alt')
            keyboard.press_and_release('m')
            keyboard.press_and_release('b')
            if wait_windows(WINDOW_NAME_PROPERTY, time_check_second=2):
                break

def del_all_group():
    pyautogui.click(150, 340)
    for _ in range(20):
        keyboard.press_and_release('del')
        time.sleep(0.2)

def click_export(base_window):
    if wait_windows(base_window, time_check_second=99):
        keyboard.press_and_release('alt')
        keyboard.press_and_release('m')
        keyboard.press_and_release('k')


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
            if wait_windows(WINDOW_NAME_OPEN, time_check_second=2):
                break
    keyboard.write(file)
    time.sleep(PAUSE_SEC)
    keyboard.press_and_release('enter')
