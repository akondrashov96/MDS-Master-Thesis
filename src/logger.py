import pandas as pd
import numpy as np
import os
from datetime import datetime

log_dir = "logs/"

class logFile:

    """
    Класс для логирования
    """

    def __init__(self, 
                 log_dir: str = log_dir, 
                 mode: str = "w", 
                 operation: str = "Загрузка данных",
                 filename: str = ""):

        "В конструкторе нужно обеспечить создание файла лога с записью об инициализации"

        now = datetime.now()

        if mode == "a":

            if filename != "":

                self.log_filename = filename

            else:
                
                print("Не указано имя файла для логирования, создание нового файла")
                self.log_filename = log_dir + "log_" + now.strftime("%d-%m-%Y %H-%M-%S") + ".txt"

        else:

            self.log_filename = log_dir + "log_" + now.strftime("%d-%m-%Y %H-%M-%S") + ".txt"
        
        self.logfile = open(self.log_filename, mode, encoding = "utf-8")

        # print(os.path.abspath(self.log_filename))
        
        self.operation = operation
        # self.logfile.write("[" + now.strftime("%d-%m-%Y %H:%M:%S") + "]: Начало - " + self.operation + "\n")
        self.write_log("Начало - " + self.operation)

    def write_log(self, message: str = "", content: str = "MSG"):

        "Добавление сообщения в лог"
        
        now_string = "[" + datetime.now().strftime("%d-%m-%Y %H:%M:%S") + f"][{content}]: "

        self.logfile.write(now_string + message + "\n")

    def close(self):

        "Отдельный метод для закрытия документа лога"

        # now = datetime.now()

        # self.logfile.write("[" + now.strftime("%d-%m-%Y %H:%M:%S") + "]: Завершение - " + self.operation)

        self.write_log("Завершение - " + self.operation)
        self.logfile.close()

    # def __del__(self):

    #     "В деструкторе нужно обеспечить закрытие файла"

    #     # now = datetime.now()

    #     # self.logfile.write("[" + now.strftime("%d-%m-%Y %H:%M:%S") + "]: Завершение - " + self.operation)

    #     self.write_log("Завершение - " + self.operation)
    #     self.logfile.close()
