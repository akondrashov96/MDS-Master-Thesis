from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer, pipeline, \
    AutoModelForTokenClassification, AutoTokenizer
import numpy as np
import re

# Универсальный путь (на HuggingFace)
# path_to_model = "ai-forever/RuM2M100-1.2B" 

# Пользовательские (локальные) пути к модели и токенизатору
path_to_model_spell = "model/M2M100ForConditionalGeneration/" 
path_to_tokenizer_spell = "model/M2M100Tokenizer/"

path_to_model_NER_names = "model/bert-finetuned-ner-names-accelerate" 

# Определение моделей

## Spellchecker
model_M100_spell = M2M100ForConditionalGeneration.from_pretrained(path_to_model_spell)
tokenizer_M100_spell = M2M100Tokenizer.from_pretrained(path_to_tokenizer_spell)

## NER для имен
label_names = ['PER-NAME', 'PER-SURN', 'PER-PATR']

id2label = {i: label for i, label in enumerate(label_names)}
label2id = {v: k for k, v in id2label.items()}

model_NER = AutoModelForTokenClassification.from_pretrained(path_to_model_NER_names,
                                                               id2label=id2label,
                                                               label2id=label2id)
tokenizer_NER = AutoTokenizer.from_pretrained(path_to_model_NER_names, use_fast=True)

token_classifier = pipeline(
    "token-classification", model=model_NER, aggregation_strategy="simple", tokenizer=tokenizer_NER
)

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

    return answer[0]


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

    NER_output = token_classifier(name_tokens)

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