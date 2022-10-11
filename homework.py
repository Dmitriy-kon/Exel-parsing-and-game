import xlrd3 as xlrd
import random

PATH = r'L:\python\python_project\life_huck\1_renaming\5_exel_file\exel\rus_eng.xlsx'


def game():
    '''
    Основной модуль игры, который вводит цикл
    '''
    print('Добро пожаловать в перевод слов')
    print('На правильный ответ я даю тебе 3 попытки\nУдачи!')
    while True:
        attempt = 3                                 # Количество попыток
        dictionary = create_dict(PATH)              # Создание словаря для игры
        question = get_random_word(dictionary)      # Выдает случайно слово из ключей словоря (рус слов)
        while True:
            if attempt == 0:
                print('Твои попытки кончились')
                get_answer(dictionary, question)
                break

            answer = input(f"Введи перевод слова: {question}\n").lower()
            if answer == 'exit':
                break
            if answer == dictionary.get(question).lower():
                print('Правильно!')
                break
            elif answer != dictionary.get(question).lower():
                print('Ответ не верный')
                attempt -= 1
                print(f'У тебя осталось {attempt} попыток')
                continue
        if another_game():
            continue
        else:
            print('Спасибо за игру!')
            break


def create_dict(file: str) -> dict:
    '''
    1. Парсит exel файл
    2. Проверяет условие
    3. Создает 2 списка с рус и англ словами и сшивает их в список
    '''
    rus = []
    eng = []
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    for row_number in range(sh.nrows):
        row = sh.row_values(row_number)
        if check(row):
            continue
        rus.append(row[0])
        eng.append(row[1])
    return {i: j for i, j in zip(rus, eng)}


def check(row: list) -> Boolean:
    '''
    Проверяет первую ячейку в строке начинается ли она с 'rus'
    '''
    return row[0].startswith('r')


def get_random_word(words: dict) -> str:
    '''
    Выдает случайны ключ из словоря
    '''
    words_key = list(words.keys())  # Создает список из ключей словаря
    return random.choice(words_key)


def another_game() -> bool:
    '''
    Проверка на повтор игры
    '''
    return input('Хочешь сыграть ещё раз? (Да или нет)').lower()\
        .startswith('д')


def get_answer(dictionary: dict, question: str) -> str:
    '''
    Хочешь ли ты получить верный перевод данного слова?
    '''
    if input('Хочешь получить верный ответ? (Да или Нет)').lower().startswith('д'):
        answer = dictionary.get(question)
        print('-' * (len(answer) + 2))
        print(answer)
        print('-' * (len(answer) + 2))


def main():
    game()


if __name__ == '__main__':
    main()
