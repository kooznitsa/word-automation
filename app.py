import win32com.client
import re
import datetime
from helpers import *


class Formatting(object):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = True

    def __init__(self, input_path, output_path, orig_doc_name, final_doc_name):
        self.input_path = input_path
        self.output_path = output_path
        self.orig_doc_name = orig_doc_name
        self.doc = self.word.Documents.Open(input_path + orig_doc_name)
        self.final_doc_name = final_doc_name

    
    def show_changes(self):
        self.doc.Activate()
        self.word.ActiveDocument.TrackRevisions = True
        self.doc.ShowRevisions = 0
        

    def add_start_text(self, start_text):
        fline = self.doc.Range(0, 0)
        fline.InsertBefore(start_text)
        fline.Font.Name = 'Times New Roman'
        fline.Font.Size = 11
        fline.Font.Italic = True
        fline.Font.Underline = 2
        fline.Font.Bold = False
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

    def convert_date(self, text):
        integers = []
        for m in text.split(' '):
            if m in months:
                integers.append(months[m])
            else:
                integers.append(m)
        
        integers = ','.join(str(x) for x in integers)
        return integers

    
    def edit_header_footer(self):
        header_primary = self.word.ActiveDocument.Sections(1).Headers(win32com.client.constants.wdHeaderFooterPrimary)
        header_fp = self.word.ActiveDocument.Sections(1).Headers(win32com.client.constants.wdHeaderFooterFirstPage)

        footer_primary = self.word.ActiveDocument.Sections(1).Footers(win32com.client.constants.wdHeaderFooterPrimary)
        footer_fp = self.word.ActiveDocument.Sections(1).Footers(win32com.client.constants.wdHeaderFooterFirstPage)

        def edit_element(element):
            for key, value in replace_dict.items():
                element.Range.Text = element.Range.Text.replace(key, value)

            element.Range.Text = re.sub(r"\"(.*?)\"", r"«\1»", element.Range.Text)
            element.Range.Text = re.sub(r"\“(.*?)\”", r"«\1»", element.Range.Text)

        edit_element(header_primary)
        edit_element(header_fp)
        edit_element(footer_primary)
        edit_element(footer_fp)


    def accept_changes(self):
        self.word.ActiveDocument.Revisions.AcceptAll()
        if self.word.ActiveDocument.Comments.Count > 0:
            self.word.ActiveDocument.DeleteAllComments()


    def close_doc(self):
        self.doc.SaveAs(self.output_path + self.final_doc_name)
        self.doc.Close()