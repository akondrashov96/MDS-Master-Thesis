# MDS-Master-Thesis
Master Thesis for MDS HSE program 

Магистерская диссертация MDS HSE

Данная диссертация подготовлена в виде проекта по обработке данных в форме с персональными данными сотрудников.

В рамках диссертации стояли основные задачи - обеспечить корректное форматирование содержимного ячеек, а также исправление орфографии.

Для выполненния данных задач мы используем как методы, основанные на правилах (через регулярные выражения), так и модели глубокого обучения (для классификации слов в именах и адресах, исправления орфографии). Для исправления орфографии используется одна из передовых моделей: [M2M100-1.2B](https://huggingface.co/ai-forever/RuM2M100-1.2B), для классификации элементов имен и адресов мы используем две отдельные модели: [ruBert-base](https://huggingface.co/ai-forever/ruBert-base?text=%D0%9C%D0%B5%D0%BD%D1%8F+%D0%B7%D0%BE%D0%B2%D1%83%D1%82+%5BMASK%5D+%D0%B8+%D1%8F+%D0%B8%D0%BD%D0%B6%D0%B5%D0%BD%D0%B5%D1%80+%D0%B6%D0%B8%D0%B2%D1%83%D1%89%D0%B8%D0%B9+%D0%B2+%D0%9D%D1%8C%D1%8E-%D0%99%D0%BE%D1%80%D0%BA%D0%B5.) для имен и [rubert-base-cased](https://huggingface.co/DeepPavlov/rubert-base-cased) для адресов.

В результате получен ETL, сканирующий входящую папку на наличие новых документов, обрабатывающий заполненную информацию и записывающий исправленные данные в шаблон формы.