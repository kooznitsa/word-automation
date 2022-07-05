import win32com.client
import re
import datetime
from helpers import *


class Formatting(object):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')

    def __init__(self, path, orig_doc_name, final_doc_name):
        self.path = path
        self.orig_doc_name = orig_doc_name
        self.doc = self.word.Documents.Open(path + orig_doc_name)
        self.final_doc_name = final_doc_name


    def add_start_text(self, start_text):
        fline = self.doc.Range(0, 0)
        fline.InsertBefore(start_text)
        fline.Font.Name = 'Times New Roman'
        fline.Font.Size = 11
        fline.Font.Italic = True
        fline.Paragraphs.Alignment = win32com.client.constants.wdAlignParagraphRight


    def replace_text(self):
        wdFindContinue = 1
        wdReplaceAll = 2

        for key, value in replace_dict.items():
            _ = self.word.Selection.Find.Execute(
                FindText=key, 
                MatchCase=False, 
                MatchWholeWord=False,
                MatchWildcards=False, 
                MatchSoundsLike=False,
                MatchAllWordForms=False, 
                Forward=True,
                Wrap=wdFindContinue, 
                Format=False, 
                ReplaceWith=value,
                Replace=wdReplaceAll
            )


    def replace_regex(self, old_regex, new_regex):
        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(old_regex, current_text):
                paragraph.Range.Text = re.sub(old_regex, new_regex, current_text)


    def edit_dates(self, regexp, symbol):
        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(regexp, current_text):
                paragraph.Range.Text = re.search(regexp, current_text).group().replace(symbol, '.')


    def convert_date(self, text):
        integers = []
        for m in text.split(' '):
            if m in months:
                integers.append(months[m])
            else:
                integers.append(m)
        
        integers = ','.join(str(x) for x in integers)
        return integers

    def format_dates(self):
        long_re = r"\d{1,2} (?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря) \d{4}"

        for p in range(1, self.doc.Paragraphs.Count):
            paragraph = self.doc.Paragraphs(p)
            current_text = paragraph.Range.Text

            if re.search(long_re, current_text):
                match = re.search(long_re, current_text).group()
                parse_date = datetime.datetime.strptime(self.convert_date(match), "%d,%m,%Y")
                new_date = parse_date.strftime("%d.%m.%Y")
                paragraph.Range.Text = re.sub(match, new_date, current_text)

    def close_doc(self):
        self.doc.SaveAs(self.path + self.final_doc_name)
        self.doc.Close()