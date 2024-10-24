import re
import time

import keyboard
import pygetwindow as pg


def wait_windows(name_like: str, time_check_second=5):
    is_win_activate = False
    max_sec = time_check_second * 10
    n = 0
    while is_win_activate is False:
        if n > max_sec:
            print('Time')
            break
        try:
            if re.search(name_like, pg.getActiveWindow().title):
                return True
        except AttributeError:
            pass
        time.sleep(0.1)
        n += 1

        if n % 10 == 0:
            print(pg.getAllTitles())
            print(f'Wait windows like [{name_like}]')

        for windows in pg.getAllTitles():
            if re.search(name_like, windows):
                is_win_activate = True
                try:
                    pg.getWindowsWithTitle(windows)[0].activate()
                except Exception:
                    continue
                time.sleep(0.3)
                return True

    return False


def window_fullscrin():
    keyboard.press_and_release('win + up')
