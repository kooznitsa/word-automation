from app import Formatting
import time
import os


input_folder = 'input_folder/'
output_folder = 'output_folder/'
path = 'C:/Users/Julia/Documents/p/code/win32com/documents/'
input_path = path + input_folder
output_path = path +output_folder
docs = next(os.walk(input_path))[2]


t1 = time.perf_counter()


def edit_docs(doc_name):
    start_text_en_ru = 'Перевод с английского языка на русский язык\n'
    start_text_fr_ru = 'Перевод с французского языка на русский язык\n'
    start_text_other = 'Перевод с английского и итальянского языков на русский язык\n'

    month_re = r"\d{1,2} (?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря) \d{4}"
    date_symbol = '.'
    date_format = '%d.%m.%Y'
    date_with_dash = r'\d{2}[-]\d{2}[-]\d{4}'
    date_with_slash = r'\d{2}[/]\d{2}[/]\d{4}'

    d = Formatting(input_path=input_path,
                            output_path=output_path,
                            doc_name=doc_name,
                            old_regex=r"\"(.*?)\"",
                            new_regex=r"«\1»")
    d.show_changes()
    d.add_start_text(start_text_en_ru)
    d.replace_text()
    d.replace_regex()
    d.edit_dates(regexp=date_with_dash, old_symbol='-', new_symbol=date_symbol)
    d.format_dates(month_re, date_format, date_symbol)
    d.edit_header_footer()
    print(d.doc_name)
    d.close_doc()


if __name__ == '__main__':
    for doc_name in docs:
        edit_docs(doc_name)
    Formatting.word.Quit()

    t2 = time.perf_counter()
    print(f'Finished in {round(t2-t1, 2)} seconds')