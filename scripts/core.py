"""
Открыть окно выбора папки и по нажатию кнопки начинает выполнять те или иные процессы
Существующие процессы:

1. Генерация АТП
2. Генерация АВР
3. Генерация АТП и АВР
4. Изменить путь к рабочей папке

Юз кейс 1:
    1. Единажды выбрать путь к папке
    2. Открыт эту папку
    3. Положить заказ HTML (и смету если он нужен)
    4. Нажать кнопку одну из вариантов "Генерировать отчет"
    5. Программма все сам дальше сделает

Юз кейс 22:
    1. Открыт эту папку
    2. Положить заказ HTML
    3. Забыть положить смету и нажать на кнопку
    4. Увидеть ошибку что нет сметы
    5. Нажать кнопку "Ок"
    6. Положить смету рядом с заказом
    7. Нажать кнопку одну из вариантов "Генерировать отчет"
    8. Программма все сам дальше сделает
"""

import os
import tkinter as tk
from tkinter import filedialog
import traceback
from scripts.models import Project
from itertools import cycle
import json
from scripts.operations import browse_folder, create_files, get_work_folder, send_message, set_work_folder, get_orders, \
    get_have_smeta

CONFIG_PATH = 'config/config.json'

def load_config():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as file:
        return json.load(file)

def save_config(config):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as file:
        json.dump(config, file, ensure_ascii=False, indent=4)


def run_project(*args, **kwargs) -> None:
    project = Project()

    root = tk.Tk()
    root.title(project.title)

    config = load_config()
    print(config)

    def update_service(value):
        config['with_service'] = value
        save_config(config)
        print(f"WITH_SERVICE updated to: {config['with_service']}")

    # 1. Генерировать АТП
    button_generate1 = tk.Button(root, text="Генерировать АТП", command=lambda: generateX("atp", project))

    # Разделитель
    label_x1 = tk.Label(root, text="")

    # 4. Генерация АТП и АВР с сметой
    folder4_var = tk.StringVar()
    label_folder4 = tk.Label(root, text="Сменить рабочую папку:")
    entry_folder4 = tk.Entry(root, textvariable=folder4_var, state="normal", width=60)
    button_folder4 = tk.Button(root, text="Выбрать", command=lambda: browse_folder(folder4_var))

    service_var = tk.BooleanVar(value=config['with_service'])
    label_service = tk.Label(root, text="Сгенерировать с услугой?")
    radio_yes = tk.Radiobutton(root, text="Да", variable=service_var, value=True, command=lambda: update_service(True))
    radio_no = tk.Radiobutton(root, text="Нет", variable=service_var, value=False, command=lambda: update_service(False))
    print(config)
    # Позиции элементов 1
    button_generate1.grid(row=0, column=0, columnspan=3, pady=10)

    # Позиция разделителя 1
    label_x1.grid(row=1, column=0, padx=10, pady=5, sticky="w")

    # Позиции элементов 4
    label_folder4.grid(row=14, column=0, padx=10, pady=5, sticky="w")
    entry_folder4.grid(row=14, column=1, padx=10, pady=5, sticky="w")
    button_folder4.grid(row=14, column=2, padx=10, pady=5)

    # Позиции радиокнопок
    label_service.grid(row=15, column=0, padx=10, pady=5, sticky="w")
    radio_yes.grid(row=15, column=1, padx=10, pady=5, sticky="w")
    radio_no.grid(row=15, column=2, padx=10, pady=5, sticky="w")

    folder = get_work_folder()
    folder4_var.set(value=f'{folder}')

    root.mainloop()


import requests
import datetime


def send_report(text=None, process=None, responsible=None):
    requests.post(
        f"https://script.google.com/macros/s/AKfycbzDwjE6Pu1a7otho2EHwbI-4yNoEmLijTfwWfI3toWpDpJ6rc-O1pKljV6XMLJmQIyJ/exec?time={datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}&process={process}&responsible={responsible}&text={text}")


def generateX(tmpl_type: str, project):
    try:
        generate(tmpl_type, project)
    except:
        if "PermissionError" in traceback.format_exc():
            text = traceback.format_exc()

            try:
                file_path = ""
                for i in text.split("\n"):
                    if "PermissionError" in i:
                        file_path = i.split("'")[1].split("/")[-1]
                        break
                send_message("Закройте файл: '" + file_path + "' и попробуйте снова")
            except:
                send_message("Неизвестная ошибка в скрипте\nОписание ошибки: " + traceback.format_exc())
        else:
            send_message("Неизвестная ошибка в скрипте\nОписание ошибки: " + traceback.format_exc())


def generate(tmpl_type: str, project):
    if project.show_warning:
        send_message(
            "В ходе работы скрипта не не открывайте/изменяйте/удаляйте файлы внутри папки так как это может привести к ошибкам\nПожалуйста дождитесь уведомления от скрипта")
    orders: dict = get_orders(project)

    if orders['status'] == -1:
        return

    for index, order in enumerate(orders['result']):

        if len(orders['result']) == 1:
            index = ""
        else:
            index = index + 1

        have_smeta: bool = get_have_smeta(order)  # type: ignore

        if "atp" == tmpl_type:  # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta,
                         index=index)  # type: ignore
        elif "avr" == tmpl_type:  # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta,
                         index=index)  # type: ignore

        elif "atp avr" == tmpl_type:  # type: ignore
            create_files(data=order, folder=get_work_folder(), tmpl_type=tmpl_type, have_smeta=have_smeta,
                         index=index)  # type: ignore

    send_report(text="АТП АВР Генератор", process="АТП АВР Генератор", responsible=os.getlogin())
    send_message("Готово!")
