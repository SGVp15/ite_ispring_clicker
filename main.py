import os.path
import re
import time

import pyautogui
import pygetwindow as pg
import pyperclip
import subprocess

from config import INFO_TICKET_IMPORT, PAUSE_SEC, BASE_PATH, EXAMS, ISPRINGQUIZMAKER_PATH


def click_ispring_import(file, base_window):
    while True:
        if wait_windows(base_window, time_second=99999):
            pyautogui.hotkey('alt')
            pyautogui.hotkey('m')
            pyautogui.hotkey('j')
            if wait_windows('Откр', time_second=2):
                break
    pyautogui.write(file)
    time.sleep(PAUSE_SEC)
    pyautogui.hotkey('enter')


def read_txt_file(path) -> ([], []):
    with open(os.path.join(path, INFO_TICKET_IMPORT), 'r', encoding='utf-8') as f:
        s = f.read()

    rows = s.split('\n')
    categories_list = []
    max_num_list = []
    num_list = []
    for row in rows:
        list_word = row.split('\t')
        if len(list_word) < 2:
            continue
        list_word = [x for x in list_word if x not in (None, '')]
        categories_list.append(list_word[0])
        num_list.append(int(list_word[1]))
        max_num_list.append(int(list_word[2]))

    files = [os.path.join(path, f'{n[:2]}.xlsx') for n in categories_list]
    return categories_list, files, num_list, max_num_list


def click_property():
    pyautogui.hotkey('alt')
    pyautogui.hotkey('m')
    pyautogui.hotkey('b')


def click_num(num):
    pyautogui.click(1600, 299)
    pyautogui.hotkey('tab')
    for n in range(num + 1):
        pyautogui.hotkey('down')


def main(path, name_window):
    categories_list, files, num_list, max_num_list = read_txt_file(path)

    for i, category in enumerate(categories_list):

        category = categories_list[i]
        file = files[i]
        num = num_list[i]
        max_num = max_num_list[i]
        if wait_windows(name_window, time_second=99999):
            click_ispring_import(file, name_window)
        else:
            return False

        time.sleep(PAUSE_SEC)
        pyperclip.copy(category)
        for _ in range(2):
            pyautogui.hotkey('shift', 'tab')
            time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(PAUSE_SEC)

        pyautogui.hotkey('enter')
        wait_windows('Результат импорта', time_second=999)
        time.sleep(PAUSE_SEC)
        pyautogui.hotkey('enter')
        time.sleep(PAUSE_SEC)

        if num != max_num:
            click_num(num)

    click_property()
    return True


def wait_windows(name_like: str, time_second=5):
    is_win_activate = False
    max_sec = time_second * 10
    n = 0
    while is_win_activate is False:
        if n > max_sec:
            print('Time')
            break

        if re.search(name_like, pg.getActiveWindow().title):
            return True

        time.sleep(0.1)
        n += 1
        for windows in pg.getAllTitles():
            if re.search(name_like, windows):
                is_win_activate = True
                try:
                    pg.getWindowsWithTitle(windows)[0].activate()
                except Exception:
                    continue
                time.sleep(0.3)
                return True
        print(f'Wait windows like [{name_like}]\t\t\t\t', end='\r')
    return False


if __name__ == '__main__':
    for exam in EXAMS:
        folders = [f for f in os.listdir(os.path.join(BASE_PATH, exam)) if
                   os.path.isdir(os.path.join(os.path.join(BASE_PATH, exam), f))]
        for num in folders:
            full_path = os.path.join(BASE_PATH, exam, num)
            subprocess.Popen([ISPRINGQUIZMAKER_PATH, os.path.join(BASE_PATH, exam, f'{exam}_{num}.quiz')],
                             creationflags=subprocess.CREATE_NEW_CONSOLE)
            name_window = f'{exam}_{num} - iSpring QuizMaker'

            if main(full_path, name_window):
                file_path_txt = os.path.join(full_path, INFO_TICKET_IMPORT)
                subprocess.Popen(["notepad", file_path_txt])
                # time.sleep(1)

# ['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',
# ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
# '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`',
# 'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',
# 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~',
# 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
# 'browserb
# ack', 'browserfavorites', 'browserforward', 'browserhome',
# 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
# 'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
# 'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
# 'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
# 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
# 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
# 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
# 'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
# 'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
# 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
# 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
# 'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
# 'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
# 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
# 'command', 'option', 'optionleft', 'optionright']
