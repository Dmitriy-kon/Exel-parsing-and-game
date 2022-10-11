import xlrd3 as xlrd


FILE = r'exel\ex_file.xlsx'


def parsing(file: str):
    book = xlrd.open_workbook(file)                                     # Открывается книга
    sh = book.sheet_by_index(0)                                         # Выбирается 1 страница из файла
    for row_number in range(sh.nrows):                                  # nrow количество строк в листе
        for col_number in range(sh.ncols):                              # ncols количество столбцов в листе
            print(sh.cell_value(rowx=row_number, colx=col_number))      # Выводит значение в ячейке по rowx (номер строки) и colx (номер столбца)


def do(file: str):
    '''
    Исполняемая функция
    1. Парсит exel файл
    2. Проверяет строку
    3. Добавляем названия и цены в списки
    '''
    articles = []
    expenses = []
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    for row_number in range(sh.nrows):
        row = sh.row_values(row_number)  # Возвращает срез значений ячеек в заданой строке
        right_row = check(row)           # Проверка списка на условие
        if right_row is None:            # Если строка таблицы не прошла условие, то continue
            continue
        articles.append(right_row[0])    #
        expenses.append(right_row[1])
    print_result(articles, expenses)


def check(row: list):
    '''
    Проверка списка содержит итого и число, то список возвращается иначе none
    '''
    if row[1]:
        if row[0] != 'Итого:' and row[1].split()[0].isdigit():
            return row


def print_result(articles: list, expenses: list):
    '''
    Принт результата
    '''
    expenses = [int(i.split()[0]) for i in expenses]    # Превращение строки вида '5000 руб' в int 5000
    index_min = get_extreme_key(expenses, min)
    index_max = get_extreme_key(expenses, max)
    print(
        f'Минимальный расход\n{articles[index_min]}\n{expenses[index_min]} рублей')
    print('-'*12)
    print(
        f'Максимальный расход\n{articles[index_max]}\n{expenses[index_max]} рублей')


def get_extreme_key(array: list, compare: str) -> int:
    '''
    Сравнивает значения согласно compare (минимальный или максимальный) и выводит индекс
    '''
    extreme_index = 0
    extreme = array[extreme_index]
    i = 1
    while i < len(array):
        if compare(array[i], extreme) == array[i]:
            extreme = array[i]
            extreme_index = i
        i += 1
    return extreme_index


def get_min_max_index(array, func):  # Более простой, но медленный вариант get_extreme_key
    element = func(array)
    return array.index(element)


if __name__ == '__main__':
    do(FILE)
