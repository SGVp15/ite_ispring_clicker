import os.path
import subprocess
import time

import keyboard

from config import INFO_TICKET_IMPORT, PAUSE_SEC, BASE_PATH, EXAMS, ISPRINGQUIZMAKER_PATH
from ispring import click_property, click_num, click_import
from windows import wait_windows, full_scrin


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


def main(path, window_name):
    categories_list, files, num_list, max_num_list = read_txt_file(path)

    for i, category in enumerate(categories_list):
        category = categories_list[i]
        file = files[i]
        num = num_list[i]
        max_num = max_num_list[i]
        if wait_windows(window_name, time_check_second=99999):
            click_import(file, window_name)
        else:
            return False

        time.sleep(PAUSE_SEC)
        for _ in range(2):
            keyboard.press_and_release('shift + tab')
            time.sleep(0.2)
        keyboard.write(category)
        time.sleep(PAUSE_SEC)
        keyboard.press_and_release('enter')

        wait_windows('Результат импорта', time_check_second=999)
        time.sleep(PAUSE_SEC)
        keyboard.press_and_release('enter')
        time.sleep(PAUSE_SEC)

        if num != max_num:
            full_scrin()
            time.sleep(0.1)
            click_num(num)

    click_property(window_name)
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

            if main(full_path, name_window):
                file_path_txt = os.path.join(full_path, INFO_TICKET_IMPORT)
                subprocess.Popen(["notepad", file_path_txt])
                # time.sleep(1)
