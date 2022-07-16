from app import Formatting
import time
import os


input_folder = 'input_folder/'
output_folder = 'output_folder/'

path = 'C:/Users/Julia/Documents/p/code/win32com/documents/'
input_path =path + input_folder
output_path =path +output_folder
docs = next(os.walk(input_path))[2]

start_text_en_ru = 'Перевод с английского языка на русский язык\n'
start_text_ru_en = 'Перевод с русского языка на английский язык\nTranslation from Russian into English\n'
start_text_fr_ru = 'Перевод с французского языка на русский язык\n'
start_text_ru_fr = 'Перевод с русского языка на французский язык\nTraduction du russe en français\n'
start_text_other = 'Перевод с английского и итальянского языков на русский язык\n'

date_with_dash = r'\d{2}[-]\d{2}[-]\d{4}'; dash_symbol = '-'
date_with_slash = r'\d{2}[/]\d{2}[/]\d{4}'; slash_symbol = '/'

t1 = time.perf_counter()


def edit_documents(doc_name):
    d = Formatting(input_path=input_path,
                            output_path=output_path,
                            orig_doc_name=doc_name,
                            final_doc_name=doc_name)
    d.show_changes()
    d.add_start_text(start_text_en_ru)
    d.replace_text()
    d.replace_regex(old_regex=r"\"(.*?)\"", new_regex=r"«\1»")
    d.replace_regex(old_regex=r"\“(.*?)\”", new_regex=r"«\1»")
    d.edit_dates(regexp=date_with_dash, symbol=dash_symbol)
    d.format_dates()
    d.edit_header_footer()
    print(d.final_doc_name)
    d.close_doc()


if __name__ == '__main__':
    for doc_name in docs:
        edit_documents(doc_name)
    Formatting.word.Quit()

    t2 = time.perf_counter()
    print(f'Finished in {round(t2-t1, 2)} seconds')