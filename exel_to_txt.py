import xlrd3 as xlrd


FILE = r'L:\python\python_project\life_huck\1_renaming\5_exel_file\exel\rus_eng.xlsx'
LEN_LINE = 20       # Длина строки в файле
DELIMITER = '_'     # Разделитель


def write_file(file: str):
    '''
    Записывает файл
    '''
    create_file(hello())
    for r, e in parse_exel(file):
        number = int(get_whitespace(r, e) / 2) * DELIMITER
        text = f'{r}{number}={number}{e}\n'
        create_file(text)


def parse_exel(file: str) -> tuple:
    '''
    1. Парсит exel файл
    2. Проверяет файл
    3. Создает 2 списка с русскими и английским словами
    4. Выводит ZIP из этих 2 списков
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
    return zip(rus, eng)


def check(row: list) -> bool:
    '''
    Проверяет начинается ли 1 ячейка в строке на 'rus'
    '''
    return row[0].startswith('r')


def hello() -> str:
    '''
    Начальная приветственная строка
    '''
    rus = 'русский'
    eng = 'английский'
    step = LEN_LINE - len(rus + eng)
    return f'{rus}{DELIMITER * step}{eng}\n{"-" * LEN_LINE}\n'


def get_whitespace(r: str, e: str) -> int:
    '''
    Считает сколько разделителей нужно добавить,
    что бы длина строки была LEN_LINE
    '''
    return LEN_LINE - len(r + e)


def create_file(inf: str):
    '''
    Записывает внутрь файла словарь строку inf
    '''
    with open('slovar.txt', 'a', encoding='utf-8') as slov:
        slov.write(inf)


def clean_file():
    '''
    Очистка файла
    '''
    with open('slovar.txt', 'w', encoding='utf-8') as slov:
        slov.write('')


def main():
    clean_file()
    write_file(FILE)


if __name__ == '__main__':
    main()
