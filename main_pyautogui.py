import pyautogui
import time
import os

# Путь к папке с чертежами
folder_path = r"D:\REP\projects\dwg_to_text\dwg"

# Список файлов DWG
files = ['organizaciya-dvizheniya-pri-dorozhnyh-rabotah.dwg']

def open_dwg(file_path):
    # Открытие AutoCAD через системный вызов (должен быть настроен путь к acad.exe)
    os.system(f'start "" "acad.exe" "{file_path}"')
    
    # Даём время для открытия файла в AutoCAD
    time.sleep(10)

def automate_cad_actions():
    # Пример автоматизации - нажатие определенных клавиш в AutoCAD
    pyautogui.press('f2')  # Открытие командной строки
    time.sleep(2)
    pyautogui.typewrite('zoom\nall\n')  # Ввод команды в AutoCAD

# Обрабатываем каждый файл
for file_name in files:
    file_path = os.path.join(folder_path, file_name)
    open_dwg(file_path)
    time.sleep(10)
    automate_cad_actions()
    time.sleep(5)  # Пауза между обработкой файлов
