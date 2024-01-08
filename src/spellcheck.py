from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer

# Универсальный путь (на HuggingFace)
# path_to_model = "ai-forever/RuM2M100-1.2B" 

# Пользовательские (локальные) пути к модели и токенизатору
path_to_model = "model/M2M100ForConditionalGeneration/" 
path_to_tokenizer = "model/M2M100Tokenizer/"

# Определение моделей
model_M100_spell = M2M100ForConditionalGeneration.from_pretrained(path_to_model)
tokenizer_M100_spell = M2M100Tokenizer.from_pretrained(path_to_tokenizer)

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