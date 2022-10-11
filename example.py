import xlrd3 as xlrd


file = r"L:\python\python_project\life_huck\1_renaming\5_exel_file\exel\ex_file.xlsx"

def parsing(file: str):
    book = xlrd.open_workbook(file)                                     # Открывается книга
    sh = book.sheet_by_index(0)                                         # Выбирается 1 страница из файла
    for row_number in range(sh.nrows):                                  # nrow количество строк в листе
        for col_number in range(sh.ncols):                              # ncols количество столбцов в листе
            print(sh.cell_value(rowx=row_number, colx=col_number))      # Выводит значение в ячейке по rowx (номер строки) и colx (номер столбца)
