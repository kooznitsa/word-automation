from app import Formatting
import time
import concurrent.futures


### ONE FILE ###
def edit_document():
    doc1 = Formatting(path='C:/Users/Julia/Documents/p/code/win32com/documents/',
                                    orig_doc_name='sample_document0.docx',
                                    final_doc_name='sample_output0.docx')

    doc1.add_start_text('Перевод с английского языка на русский язык\n')
    doc1.replace_text()
    doc1.replace_regex(old_regex=r"\"(.*?)\"", new_regex=r"«\1»")
    doc1.edit_dates(regexp=r'\d{2}[-]\d{2}[-]\d{4}', symbol='-')
    doc1.format_dates()
    doc1.close_doc()


### SEVERAL FILES ###
docs = ['sample_document0.docx', 'sample_document1.docx', \
    'sample_document2.docx']

t1 = time.perf_counter()

def edit_documents(doc_path):
    d = Formatting(path='C:/Users/Julia/Documents/p/code/win32com/documents/',
                            orig_doc_name=doc_path,
                            final_doc_name=f'sample_document{[docs.index(doc_path)]}_output.docx')
    d.add_start_text('Перевод с английского языка на русский язык\n')
    d.replace_text()
    d.replace_regex(old_regex=r"\"(.*?)\"", new_regex=r"«\1»")
    d.edit_dates(regexp=r'\d{2}[-]\d{2}[-]\d{4}', symbol='-')
    d.format_dates()
    print(d.final_doc_name)
    d.close_doc()

if __name__ == '__main__':
    with concurrent.futures.ProcessPoolExecutor() as executor:
        executor.map(edit_documents, docs)
    Formatting.word.Quit()

    t2 = time.perf_counter()

    print(f'Finished in {round(t2-t1, 0)} seconds')