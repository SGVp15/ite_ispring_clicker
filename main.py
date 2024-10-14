import os.path
import re
import subprocess
import time

import keyboard
import pyautogui
import pygetwindow as pg

from config import INFO_TICKET_IMPORT, PAUSE_SEC, BASE_PATH, EXAMS, ISPRINGQUIZMAKER_PATH


def click_ispring_import(file, base_window):
    while True:
        if wait_windows(base_window, time_second=99999):
            keyboard.press_and_release('alt')
            keyboard.press_and_release('m')
            keyboard.press_and_release('j')
            if wait_windows('Откр', time_second=2):
                break
    keyboard.write(file)
    time.sleep(PAUSE_SEC)
    keyboard.press_and_release('enter')


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
    keyboard.press_and_release('alt')
    keyboard.hotkey('m')
    keyboard.press_and_release('b')


def click_num(num):
    pyautogui.click(1600, 299)
    keyboard.press_and_release('tab')
    des = num // 10

    keyboard.press_and_release(str(des))
    keyboard.press_and_release(str(des))

    for n in range(num % 10):
        keyboard.press_and_release('down')


def main(path, window_name):
    categories_list, files, num_list, max_num_list = read_txt_file(path)

    for i, category in enumerate(categories_list):

        category = categories_list[i]
        file = files[i]
        num = num_list[i]
        max_num = max_num_list[i]
        if wait_windows(window_name, time_second=99999):

            click_ispring_import(file, window_name)
        else:
            return False

        time.sleep(PAUSE_SEC)
        # pyperclip.copy(category)
        for _ in range(2):
            keyboard.press_and_release('shift + tab')
            time.sleep(0.1)
        keyboard.write(category)
        # keyboard.press_and_release('ctrl + v')
        time.sleep(PAUSE_SEC)

        keyboard.press_and_release('enter')
        wait_windows('Результат импорта', time_second=999)
        time.sleep(PAUSE_SEC)
        keyboard.press_and_release('enter')
        time.sleep(PAUSE_SEC)

        if num != max_num:
            keyboard.press_and_release('win + up')
            time.sleep(0.1)
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
