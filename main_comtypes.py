import comtypes.client
import os
import time
import icecream

# Путь к папке с чертежами
folder_path = r"D:\REP\projects\dwg_to_text\dwg"

# Список файлов DWG
files = ['organizaciya-dvizheniya-pri-dorozhnyh-rabotah.dwg']

def process_dwg(file_path):
    # Создаем COM объект AutoCAD
    acad = comtypes.client.CreateObject("AutoCAD.Application")
    acad.Visible = True

    try:
        # Открываем чертеж
        doc = acad.Documents.Open(file_path)
        time.sleep(10)  # Даем время AutoCAD загрузить чертеж

        # Получаем пространство модели
        model_space = doc.ModelSpace

        # Множество для хранения уникальных имен объектов
        object_names = set()

        # Списки для хранения текста и мультитекста
        text_entities = []
        mtext_entities = []

        # Итерируемся по всем объектам в пространстве модели
        for entity in model_space:
            try:
                # Добавляем имя объекта в множество
                object_names.add(entity.ObjectName)
            except Exception as entity_error:
                print(f"Ошибка при обработке объекта: {str(entity_error)}")

            try:
                # Проверяем тип объекта на MTEXT
                if entity.ObjectName == "AcDbMText":
                    mtext_entities.append(entity.TextString)
            except Exception as entity_error:
                print(f"Ошибка при обработке текста объекта: {str(entity_error)}")

        # Выводим результаты для текущего чертежа
        print(f"Чертеж: {file_name}")
        print("\nMTEXT объекты:")
        icecream.ic(mtext_entities)

        print("\n" + "-" * 50 + "\n")

        # Выводим список уникальных ObjectName
        print(f"Чертеж: {file_path}")
        print("Уникальные ObjectName в DWG файле:")
        icecream.ic(object_names)

        print("\n" + "-" * 50 + "\n")

        # Закрываем документ
        doc.Close(False)
        time.sleep(1)  # Даем время AutoCAD закрыть файл

    except Exception as file_error:
        print(f"Ошибка при обработке {file_path}: {str(file_error)}")

    finally:
        try:
            # Закрываем AutoCAD
            acad.Quit()
        except Exception as quit_error:
            print(f"Ошибка при закрытии AutoCAD: {str(quit_error)}")

# Обрабатываем каждый файл из списка
for file_name in files:
    file_path = os.path.join(folder_path, file_name)
    process_dwg(file_path)
    time.sleep(5)  # Задержка между файлами для стабилизации AutoCAD
