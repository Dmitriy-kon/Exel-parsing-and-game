import xlrd3 as xlrd


FILE = r'exel\ex_file.xlsx'


def parsing(file):
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    for row_number in range(sh.nrows):
        for col_number in range(sh.ncols):
            print(sh.cell_value(rowx=row_number, colx=col_number))


def do(file):
    articles = []
    expenses = []
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    for row_number in range(sh.nrows):
        row = sh.row_values(row_number)
        right_row = check(row)
        if right_row is None:
            continue
        articles.append(right_row[0])
        expenses.append(right_row[1])
    print_result(articles, expenses)


def check(row):
    if row[1]:
        if row[0] != 'Итого:' and row[1].split()[0].isdigit():
            return row


def print_result(articles, expenses):
    expenses = [int(i.split()[0]) for i in expenses]
    index_min = get_extreme_key(expenses, min)
    index_max = get_extreme_key(expenses, max)
    # index_min = get_min_max_index(expenses, min)
    # index_max = get_min_max_index(expenses, max)
    print(
        f'Минимальный расход\n{articles[index_min]}\n{expenses[index_min]} рублей')
    print('-'*12)
    print(
        f'Максимальный расход\n{articles[index_max]}\n{expenses[index_max]} рублей')


def get_extreme_key(array, compare):
    extreme_index = 0
    extreme = array[extreme_index]
    i = 1
    while i < len(array):
        if compare(array[i], extreme) == array[i]:
            extreme = array[i]
            extreme_index = i
        i += 1
    return extreme_index


def get_min_max_index(array, func):
    element = func(array)
    return array.index(element)


if __name__ == '__main__':
    # parsing(FILE)
    do(FILE)
    # print('Время работы функции', timeit.timeit(
    #     "do(FILE)", setup='from __main__ import FILE, do', number=10))
