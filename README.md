# python_docx_to_xlsx

find_docx_files:  обходит указанную директорию и находит все файлы с расширением .docx, добавляя их в список.

read_docx_lines: считывает строки из каждого файла .docx и возвращает их в виде списка.

write_to_excel: создает новый файл Excel и записывает строки из всех найденных файлов .docx в отдельные строки. 
Каждая строка из документа .docx будет записана в одну строку файла .xlsx, а каждая ячейка в этой строке будет 
содержать одну строку из документа.

main: Основная функция, которая запрашивает у пользователя путь к директории и запускает процесс.
