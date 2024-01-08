import pandas as pd
import numpy as np
import os
import re
import locale
import pymorphy3
from dateutil import parser
from datetime import datetime

from src.logger import logFile
from src.utilities import RussianParserInfo
from src.spellcheck import correct_errors


# Настройки
workdir = "data/"
locale.setlocale(locale.LC_ALL, 'ru_RU')
m = pymorphy3.MorphAnalyzer()

def correct_date_condition2(datestring: str) -> str:

    """
    Функция для проверки и корректировки по условию 2: 
    Графы «Поступление» и «Увольнение» пункта 14 даты должны содержать только цифры и точки.

    Функция проверяет соответствие даты формату ДД.ММ.ГГГГ или ММ.ГГГГ.
    Функция угадывает формат даты, если указана дата в читаемом формате и возвращает строку в формате ММ.ГГГГ.
    В противном случае функция оставляет в строке только числа и точки. 
    В случае, если указано "по настоящее время" - текст не меняется.

    Параметры:
    datestring : str
        Дата для форматирования

    Возвращает:
    string_out : str
        Строка с датой в формате ДД.ГГГГ

    """

    if datestring.lower() == "по настоящее время": return datestring
    
    try:

        dt = datetime.strptime(datestring, "%m.%Y")
        return(datestring)
    
    except ValueError as e:

        try:

            dt = datetime.strptime(datestring, "%d.%m.%Y")
            return(dt.strftime("%m.%Y"))
        
        except ValueError as e:

            try: 
                
                return(date_guesser_corrector(datestring, "%m.%Y"))
            
            except Exception as e:

                # нечитаемый формат даты
                pass

    return("".join(re.findall(pattern = "[0-9.]+", string = datestring)))

def only_cyrillic(textstring: str) -> str:

    """
    Функция для проверки и корректировки по условию 3.1: 
    1. Графа «Степень родства» пункта 15 должна содержать только буквы кириллицы
    2. В графе «Фамилия, имя и отчество» предыдущие фамилии (девичьи, изменённые) указываются в скобках. 
    Если у одного родственника несколько раз изменялась фамилия, то они указываются в скобках через запятую.

    Функция оставляет в строке только числа и точки. В случае, если указано "по настоящее время" - текст не меняется.

    Параметры:
    datestring : str
        Дата для форматирования

    Возвращает:
    string_out : str
        Строка, содержащая слово в желаемом падеже

    """

    return("".join(re.findall(pattern = "[а-яА-Я ]+", string = textstring)))

def last_surnames():
    
    pass



def case_corrector(string_in: str, word: str, case: str = "nomn") -> str:

    """
    Функция для исправления падежа слова в предложении. 
    Делается через подстановку заполняемого выражения в строку и форматирование.

    Параметры:
    string_in : str
        Строка, в которой есть слово для форматирования
    word : str
        Слово, которое необходимо скорректировать (используется как паттерн)
    case : str
        Падеж, в который необходимо просклонять слово

    Возвращает:
    string_out : str
        Строка, содержащая слово в желаемом падеже
    """

    string_out = re.sub(word, "{replacement}", string_in)\
            .format(replacement = m.parse(word)[0].inflect({case}).word.title())
    
    return string_out


def date_guesser_corrector(datestring: str, target_format: str = "%Y, %d %B") -> str:

    """
    Функция для определения datetime из произвольного формата даты и конвертации полученного
    datetime в указанный формат.

    Параметры:
    datestring : str
        Строка, содержащая дату в произвольном формате (на русском)
    target_format : str
        Желаемый формат даты

    Возвращает:
    datestring_corrected : str
        Строка, содержащая дату в желаемом формате
    """

    # Создание паттернов для поиска форматов и частей даты
    date_pattern = "\%[a-zA-Z]"
    pattern_split = "[0-9а-яА-Я]+"

    formats = re.findall(date_pattern, target_format)

    # Угадывание даты из строки
    guessdate = parser.parse(datestring, 
                             parserinfo = RussianParserInfo())                         
    
    # если в искомом формате присутствует месяц, необходимо применить верное склонение
    if "%B" in formats: 
        
        guessdate_inner = guessdate.strftime("%d %B %Y")
        dateparts_inner = re.findall(pattern_split, guessdate_inner)

        datestring_corrected = case_corrector(guessdate.strftime(target_format), dateparts_inner[1], 'gent')
    
    else:

        datestring_corrected = guessdate.strftime(target_format)
    
    return datestring_corrected


def pages_preprocessor(filename: str, workdir: str = workdir, logfile: str = "", verbose: bool = False):

    """
    Чтение и предобработка данных для каждого листа анкеты. 

    На листе 1 проверяется условие 1: 
        В пункте 3 год рождения указывается только цифрами, число – двумя цифрами
    На листе 2 проверяется условие 2: 
        Графы «Поступление» и «Увольнение» пункта 14 даты должны содержать только цифры и точки.
    На листе 3 проверяются условия 3.1 и 3.2: 
        1. Графа «Степень родства» пункта 15 должна содержать только буквы кириллицы.
        2. В графе «Фамилия, имя и отчество» предыдущие фамилии (девичьи, изменённые) указываются в скобках. 
        Если у одного родственника несколько раз изменялась фамилия, то они указываются в скобках через запятую.
    На листах 2-4 проверяется условие 4: 
        При заполнении адресов проживания и работы сначала необходимо указывать регион: республику, край, область.
    
    Параметры:
    filename : str
        Имя файла
    workdir : str
        Рабочая директория
    """

    if verbose: print(f"Обработка документа {filename}")

    # Начать логирование: лог 1
    if logfile == "" and len(os.listdir("logs")) != 0:

        logfile = os.listdir("logs")[-1]

    log = logFile(mode = "a", filename = "logs/"+logfile, operation = "Обработка")

    data_sheets = [pd.read_excel(workdir + "raw/" + filename,
                                 sheet_name=sheet_index) for sheet_index in range(4)]
    
    # Лог 2
    msg = f"{filename}: Листы считаны"
    log.write_log(msg)

    if verbose: print(msg)

    
    # Лист 1: обработка

    # Лог 3
    msg = f"{filename}: Обработка первого листа"
    log.write_log(msg)

    # Переименование колонок и удаление пустот
    data_sheets[0].columns = ["Вопрос", "Ответ"]
    data_sheets[0] = data_sheets[0][2:].dropna(subset = ["Вопрос"]).reset_index(drop = True)

    # Проверка орфографии
    print("Проверка орфографии...")
    data_sheets[0]["Ответ"] = data_sheets[0]["Ответ"].map(correct_errors)

    print("Завершена!")
    
    # Лог 4
    msg = f"{filename}: Лист 1: орфография проверена"
    log.write_log(msg)

    # Проверка условия 1: В пункте 3 год рождения указывается только цифрами, число – двумя цифрами 
    date_place_parts = data_sheets[0]['Ответ'][4].split(", ")

    # Проверка на вхождение паттерна
    regex_matcher = re.search(pattern = r"\d{4},\s\d{2}\s\w+", 
                              string = date_place_parts[0] + ", " + date_place_parts[1])
    
    if regex_matcher is None:

        # Вычленение даты произвольного формата и переформативание в ГГГГ, ДД ММММ
        try:

            date_corrected = date_guesser_corrector(date_place_parts[0] + ", " + date_place_parts[1])

            date_place_corrected = " ".join([date_corrected, 
                                             *date_place_parts[2:]])
            
            data_sheets[0]['Ответ'][4] = date_place_corrected

        except Exception as e:
            
            # Лог error_1
            msg = f"{filename}: Лист 1: ошибка в формате даты"
            log.write_log(msg, content = "ERR")
            log.write_log(str(e), content = "ERR")

    else:

        # Нужно ли исправить падеж в дате при вхождении паттерна?
        pass

    # Лог 5
    msg = f"{filename}: Лист 1: Условие 1 исправлено"
    log.write_log(msg)


    # Лист 2: обработка

    # Лог 6
    msg = f"{filename}: Обработка второго листа"
    log.write_log(msg)

    data_sheets[1].columns = ["Месяц и год поступления", 
                              "Месяц и год увольнения",
                              "Должность с указанием наименования организации",
                              "Адрес организации"]

    data_sheets[1] = data_sheets[1][2:].dropna().reset_index(drop = True)

    # Проверка условия 2: Графы «Поступление» и «Увольнение» пункта 14 даты должны содержать только цифры и точки.
    data_sheets[1]["Месяц и год поступления"] = data_sheets[1]["Месяц и год поступления"].map(correct_date_condition2)
    data_sheets[1]["Месяц и год увольнения"] = data_sheets[1]["Месяц и год увольнения"].map(correct_date_condition2)

    # Лог 7
    msg = f"{filename}: Лист 2: Условие 2 исправлено"
    log.write_log(msg)

    # Проверка орфографии
    print("Проверка орфографии...")
    data_sheets[1]["Должность с указанием наименования организации"] = data_sheets[1]["Должность с указанием наименования организации"].map(correct_errors)
    data_sheets[1]["Адрес организации"] = data_sheets[1]["Адрес организации"].map(correct_errors)

    print("Завершена!")
    
    # Лог 8
    msg = f"{filename}: Лист 1: орфография проверена"
    log.write_log(msg)
    

    # # Обработка третьего листа
    # data_sheets[2].columns = ["Степень родства", 
    #                       "Фамилия, имя и отчество",
    #                       "Число, месяц, год и место рождения, гражданство",
    #                       "Место работы, должность",
    #                       "Адрес места жительства"]

    # data_sheets[2] = data_sheets[2][1:].dropna().reset_index(drop = True)


    # Обработка четвертого листа
    
