import os.path
import subprocess
import time

import keyboard
import pyautogui

from config import INFO_TICKET_IMPORT, PAUSE_SEC, BASE_PATH, EXAMS, ISPRINGQUIZMAKER_PATH, \
    WINDOW_NAME_IMPORT_FROM_EXCEL, WINDOW_NAME_RESULT_IMPORT_FROM_EXCEL, WINDOW_NAME_PROPERTY, WINDOW_NAME_EXPORT
from ispring import click_property, click_num, click_import, click_export
from windows import wait_windows, window_fullscrin


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


def run_clicker(path, window_name):
    categories_list, files, num_list, max_num_list = read_txt_file(path)

    for i, category in enumerate(categories_list):
        category = categories_list[i]
        file = files[i]
        num = num_list[i]
        max_num = max_num_list[i]

        wait_windows(window_name)
        window_fullscrin()
        time.sleep(PAUSE_SEC)
        pyautogui.click(150, 340)
        for _ in range(20):
            keyboard.press_and_release('del')
            time.sleep(0.2)

        click_import(file, window_name)
        wait_windows(WINDOW_NAME_IMPORT_FROM_EXCEL, time_check_second=999)

        time.sleep(PAUSE_SEC)
        for _ in range(2):
            keyboard.press_and_release('shift + tab')
            time.sleep(0.2)
        keyboard.write(category)
        time.sleep(PAUSE_SEC)
        keyboard.press_and_release('enter')

        wait_windows(WINDOW_NAME_RESULT_IMPORT_FROM_EXCEL, time_check_second=999)
        time.sleep(PAUSE_SEC)
        keyboard.press_and_release('enter')
        time.sleep(PAUSE_SEC)

        if num != max_num:
            window_fullscrin()
            time.sleep(0.1)
            click_num(num)

    click_property(window_name)
    wait_windows(WINDOW_NAME_PROPERTY, time_check_second=999)
    time.sleep(PAUSE_SEC * 2)
    pyautogui.click(510, 256)
    time.sleep(PAUSE_SEC * 2)
    for _ in range(3):
        keyboard.press_and_release('tab')
        time.sleep(PAUSE_SEC * 2)
    keyboard.press_and_release('space')
    time.sleep(PAUSE_SEC * 2)
    keyboard.press_and_release('tab')
    time.sleep(PAUSE_SEC * 2)

    time.sleep(1)
    pyautogui.click(900, 900)
    keyboard.press_and_release('tab')
    time.sleep(PAUSE_SEC)
    keyboard.press_and_release('enter')
    time.sleep(PAUSE_SEC)
    keyboard.press_and_release('ctrl+s')
    time.sleep(2)
    click_export(window_name)
    return True


if __name__ == '__main__':
    for exam in EXAMS:
        folders = [f for f in os.listdir(os.path.join(BASE_PATH, exam)) if
                   os.path.isdir(os.path.join(os.path.join(BASE_PATH, exam), f))]
        for num in folders:
            full_path = os.path.join(BASE_PATH, exam, num)
            subprocess.Popen([ISPRINGQUIZMAKER_PATH, os.path.join(BASE_PATH, exam, f'{exam}_{num}.quiz')],
                             creationflags=subprocess.CREATE_NEW_CONSOLE)
            name_window = f'{exam}_{num} - iSpring QuizMaker'

            if run_clicker(full_path, name_window):
                file_path_txt = os.path.join(full_path, INFO_TICKET_IMPORT)
                subprocess.Popen(["notepad", file_path_txt])
