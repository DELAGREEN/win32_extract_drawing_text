import win32com.client
import os
import time
import icecream

# Путь к папке с чертежами
folder_path = r"D:\REP\projects\dwg_to_text\dwg"

# Получаем список всех файлов в папке
#files = [f for f in os.listdir(folder_path) if f.endswith('.dwg')]
files = ['organizaciya-dvizheniya-pri-dorozhnyh-rabotah.dwg']

def process_dwg(file_path):
    # Создаем COM объект AutoCAD
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True

    try:
        # Открываем чертеж
        doc = acad.Documents.Open(file_path)
        time.sleep(10)

        # Получаем пространство модели
        model_space = doc.ModelSpace

        # Множество для хранения уникальных имен объектов
        object_names = set()

        # Списки для хранения текста
        text_entities = []
        mtext_entities = []

        # Итерируемся по всем объектам в пространстве модели
        for entity in model_space:

            try:
                # Добавляем имя объекта в множество
                object_names.add(entity.ObjectName)
            except Exception as entity_error:
                print(f"Failed to process entity: {str(entity_error)}")


            try:
                # Проверяем тип объекта на TEXT
                #if entity.ObjectName == "AcDbText":
                #    text_entities.append(entity.TextString)

                # Проверяем тип объекта на MTEXT
                #elif entity.ObjectName == "AcDbMText":
                if entity.ObjectName == "AcDbMText":
                    mtext_entities.append(entity.TextString)

            except Exception as entity_error:
                print(f"Failed to process entity: {str(entity_error)}")

        # Выводим результаты для текущего файла
        print(f"File: {file_name}")
        print("\nMTEXT Entities:")
        icecream.ic(mtext_entities)

        print("\n" + "-"*50 + "\n")

        # Выводим список уникальных ObjectName
        print(f"File: {file_path}")
        print("Unique ObjectNames in the DWG file:")
        icecream.ic(object_names)

        print("\n" + "-"*50 + "\n")

        # Закрываем документ
        doc.Close(False)
        time.sleep(1)  # Даем время AutoCAD обработать закрытие

    except Exception as file_error:
        print(f"Failed to process {file_path}: {str(file_error)}")

    finally:
        try:
            acad.Quit()
        except Exception as quit_error:
            print(f"Failed to quit AutoCAD: {str(quit_error)}")

# Проходимся по каждому файлу в папке
for file_name in files:
    file_path = os.path.join(folder_path, file_name)
    process_dwg(file_path)
    time.sleep(5)  # Увеличенная задержка между файлами для стабилизации AutoCAD
