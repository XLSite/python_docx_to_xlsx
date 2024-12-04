import os
from openpyxl import Workbook
from docx import Document

def find_docx_files(directory):
    """Обходит указанную директорию и находит все файлы .docx."""
    docx_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                docx_files.append(os.path.join(root, file))
    return docx_files

def read_docx_lines(file_path):
    """Считывает строки из файла .docx."""
    document = Document(file_path)
    lines = []
    for paragraph in document.paragraphs:
        lines.append(paragraph.text)
    return lines

def write_to_excel(docx_files, output_file):
    """Записывает строки из .docx файлов в .xlsx файл."""
    workbook = Workbook()
    sheet = workbook.active

    for docx_file in docx_files:
        lines = read_docx_lines(docx_file)
        # Записываем каждую строку в отдельную ячейку текущей строки
        sheet.append(lines)  # Каждая строка из .docx будет записана в одну строку .xlsx

    workbook.save(output_file)

def main():
    directory = input("Введите путь к директории: ")
    output_file = "output.xlsx"  # Имя выходного файла

    docx_files = find_docx_files(directory)
    write_to_excel(docx_files, output_file)

    print(f"Данные из {len(docx_files)} файлов .docx записаны в {output_file}")

if __name__ == "__main__":
    main()