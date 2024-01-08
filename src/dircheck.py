import pandas as pd
import numpy as np
import os
from datetime import datetime
from src.logger import logFile

workdir = "../data/"

def get_new_file_names(dir_in: str, dir_out: str, verbose: bool = False) -> list:

    """
    Проверка, были ли изменения в файлах в папках для ввода и вывода. 
    Функция считывает только документы excel 
    Так как в имени файла ожидается наличие timestamp, то если будет обновление по анкете, это будет считано как 
    новый документ. 
    
    Параметры:
    dir_in : str
        Путь к папке с данными на обработку
    dir_out : str
        Путь к папке с обработанными файлами

    Возвращает:
    new_files : list of str
        Список имен файлов, подходящих критериям
    """

    # Начать логирование: лог 1
    log = logFile()

    # Получить список файлов
    files_in = set([filename for filename in os.listdir(dir_in) if filename.split(".")[-1] in ["xls", "xlsx"]])
    files_out = set([filename.replace("_check", "") for filename in os.listdir(dir_out)])

    # Очистить список входящих файлов от открытых Экселей (в начале названия которых есть "~$")
    for filename in list(files_in):

        if "~$" in filename:

            # Лог warning_1
            msg = f"{filename} Имеются открытые документы, исключение из выдачи"
            log.write_log(msg)

            if verbose: print(msg)

            files_in.remove(filename)
            files_in.remove(filename[2:])

    # Лог 2
    msg = "Получены имена документов в папках ввода и вывода"
    log.write_log(msg)

    # print(files_in)

    # Проверить наличие timestamp правильного формата в имени файла
    
    error_names = []

    for filename in list(files_in):

        date_string = filename.split("_")[-1].split(".")[0]

        # в timestamp должно быть 14 знаков + только числа
        if len(date_string) != 14 and date_string.isnumeric():

            msg = f"Ошибка в {filename}: Неверный формат timestamp" 

            if verbose: print(msg)

            # Лог error_1
            log.write_log(msg, content = "ERR")
            
            error_names.append(filename)
            
            next

        else:
            try:

                dt = datetime.strptime(date_string, "%d%m%Y%H%M%S")
                # print(dt)

            except ValueError as e:

                msg = f"Ошибка в {filename}: Неверный формат timestamp" 
                if verbose: print(msg)

                # Лог error_2a
                log.write_log(msg, content = "ERR")

                error_names.append(filename)

            except Exception as e:

                msg = f"Ошибка в {filename}: Неизвестная ошибка" 
                if verbose: print(msg)

                # Лог error_2b
                log.write_log(msg, content = "ERR")
                log.write_log(str(e), content = "ERR")
        
    files_in = files_in - set(error_names)

    # Определить новые файлы
    new_files = list(files_in - files_out)

    # Лог 3
    msg = "Получены имена документов для обработки\n" + "\n"\
        .join(['\t' + str(i) + ". " + filename + " timestamp: " +\
               datetime.strptime(filename.split("_")[-1].split(".")[0], "%d%m%Y%H%M%S")\
                .strftime("%d-%m-%Y %H:%M:%S") for i, filename in enumerate(new_files)])
    log.write_log(msg)

    # Закрытие лога: лог 4
    log.close()

    return new_files