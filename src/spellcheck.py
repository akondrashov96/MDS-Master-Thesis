from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer, pipeline, \
    AutoModelForTokenClassification, AutoTokenizer
import numpy as np
import re

# Универсальный путь (на HuggingFace)
# path_to_model = "ai-forever/RuM2M100-1.2B" 

# Пользовательские (локальные) пути к модели и токенизатору
path_to_model_spell = "model/M2M100ForConditionalGeneration/" 
path_to_tokenizer_spell = "model/M2M100Tokenizer/"

path_to_model_NER_names = "model/stable/bert-finetuned-ner-names-accelerate" 
path_to_model_NER_addresses = "model/stable/bert-finetuned-ner-addresses-accelerate" 

# Определение моделей

## Spellchecker
model_M100_spell = M2M100ForConditionalGeneration.from_pretrained(path_to_model_spell)
tokenizer_M100_spell = M2M100Tokenizer.from_pretrained(path_to_tokenizer_spell)

## NER для имен
label_names = ['PER-NAME', 'PER-SURN', 'PER-PATR']

id2label = {i: label for i, label in enumerate(label_names)}
label2id = {v: k for k, v in id2label.items()}

model_NER_name = AutoModelForTokenClassification.from_pretrained(path_to_model_NER_names,
                                                               id2label=id2label,
                                                               label2id=label2id)
tokenizer_NER_name = AutoTokenizer.from_pretrained(path_to_model_NER_names, use_fast=True)

token_classifier_name = pipeline(
    "token-classification", model=model_NER_name, aggregation_strategy="simple", tokenizer=tokenizer_NER_name
)

## NER для адресов
label_names = ["LOC-REG", 
               "LOC-DIST", 
               "LOC-SETL", 
               "LOC-CDIST", 
               "LOC-STRT", 
               "LOC-HOUS", 
               "LOC-FLAT"]

id2label = {i: label for i, label in enumerate(label_names)}
label2id = {v: k for k, v in id2label.items()}

model_NER_adr = AutoModelForTokenClassification.from_pretrained(path_to_model_NER_addresses,
                                                               id2label=id2label,
                                                               label2id=label2id)

tokenizer_NER_adr = AutoTokenizer.from_pretrained(path_to_model_NER_addresses, use_fast=True)

token_classifier_adr = pipeline(
    "token-classification", model=model_NER_adr, aggregation_strategy="simple", tokenizer=tokenizer_NER_adr
)


# Функции

def correct_errors(sentence_in: str) -> str:

    """
    Функция для исправления орфографии в модели 

    Параметры:
    sentence_in : str
        Предложение с (возможно) орфографическими ошибками

    Возвращает:
    answer : str
        Предложение, очищенное от ошибок
    """

    encodings = tokenizer_M100_spell(sentence_in, return_tensors="pt")
    generated_tokens = model_M100_spell.generate(**encodings, 
                                                 forced_bos_token_id=tokenizer_M100_spell.get_lang_id("ru"), 
                                                 max_new_tokens = 200)
    answer = tokenizer_M100_spell.batch_decode(generated_tokens, skip_special_tokens=True)

    return re.sub('[^а-яА-яёЁ\.\-, ]+', '', answer[0])


def name_reconstruct(name: str) -> str:

    """
    Функция для исправления формата имен в формат ФИО 
    В случае, если в тексте распознается более 1 фамилии, то используется формат 
        Ф (Ф1, Ф2, ... - при наличии старых фамилий) И О

    Параметры:
    name : str
        Строка с именем

    Возвращает:
    string_out : str
        Строка с именем требуемого формата
    """

    # создание словаря для сортировки элементов имени
    entities = ['SURN', 'NAME', 'PATR']
    sort_dict = {key: elem for elem, key in list(enumerate(entities))}

    name_tokens = re.findall("[а-яА-ЯЁё\-]+", name)

    NER_output = token_classifier_name(name_tokens)

    name_classes =  np.array([elem[0]['entity_group'] for elem in NER_output])

    # переформирование имени
    string_out = " ".join([x for _, x in sorted(zip(name_classes, name_tokens), key = lambda pair: sort_dict[pair[0]])])

    nameparts_counts = np.unique(name_classes, return_counts=True)

    # В случае, если больше одной фамилии - фамилии, следующие после 1й заключить в скобки
    if "SURN" in nameparts_counts[0] and nameparts_counts[1][-1] > 1:

        surnames = string_out.split()[:nameparts_counts[1][-1]]
        other_name_part = string_out.split()[nameparts_counts[1][-1]:]

        string_out = surnames[0] + " (" + ", ".join(surnames[1:]) + ") " + " ".join(other_name_part)

    return(string_out)


def address_reconstruct(address: str) -> str:

    """
    Функция для исправления формата адреса в формат Регион, район, город/поселок, улица, дом, квартира 

    Параметры:
    address : str
        Строка, содержащая адрес

    Возвращает:
    string_out : str
        Строка с адресом требуемого формата
    """

    # создание словаря для сортировки элементов адреса
    entities = ["O", "REG", "DIST", "SETL", "CDIST", "STRT", "HOUS", "FLAT"]
    
    sort_dict = {key: elem for elem, key in list(enumerate(entities))}

    adr_tokens = address.strip().split(", ")

    # print(adr_tokens)

    NER_output = list(token_classifier_adr(adr_tokens))

    addresses = [[]]
    adr_entities = [[]]
    i = 0

    for token, elem in zip(adr_tokens, NER_output):

        if len(elem) > 1:

            if elem[0]["start"] > 0:

                i += 1
                addresses.append([token[:elem[0]["start"]]])
                adr_entities.append(["O"])

            for subelem in elem:

                addresses[i].append(subelem["word"])
                adr_entities[i].append(subelem["entity_group"])

            if elem[-1]["end"] < len(token)-1:

                addresses.append([token[elem[-1]["end"]:]])
                adr_entities[i].append("O")

        else:
            addresses[i].append(token)
            adr_entities[i].append(elem[0]["entity_group"])

    # print(addresses)
    # print(adr_entities)

    final_list = []

    for adr_lst, entity in zip(addresses, adr_entities):      

        # print("tokens:", adr_lst)
        # print("entities:", entity)    

        if entity[0] == "O":
                idx_start = 1
                left = adr_lst[0]
        else:
                idx_start = 0
                left = ""

        if entity[-1] == "O":
                idx_end = len(entity) - 1
                right = adr_lst[-1]
        else:
                idx_end = len(entity)
                right = ""

        adr_sorted = [x for _, x in sorted(zip(entity[idx_start:idx_end], adr_lst[idx_start:idx_end]), \
                                            key = lambda pair: sort_dict[pair[0]])]

        final_list.append(left + ", ".join(adr_sorted) + right)
        
        # print(final_list)

    # # переформирование адреса
    string_out = ", ".join(final_list)

    return(string_out)