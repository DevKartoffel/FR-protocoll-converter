import csv
import unicodedata
import re
from docx import Document
from enum import Enum

class Categories(Enum):
    CHILDREN = 1
    PARENT = 2
    UNDEFINED = 3

class ClassifiedParagraph():
    FILTER_EXEPTIONS = ['child', 'parent']
    
    def __init__(self, paraText:str):
        self.category = self.classify(paraText)
        if self.category != Categories.UNDEFINED:
            self.text = paraText[2:].strip()
        else:
            self.text = paraText
        self.numberWords = self.count_words()
        self.numberLetters = self.count_chars()

    def classify(self, paraText):
        first_letters = paraText[0:3].lower()
        if 'c:' in first_letters:
            return Categories.CHILDREN
        
        if 'p:' in first_letters:
            return Categories.PARENT
        
        return Categories.UNDEFINED
    
    def get_clean_text(self):
        # Filter [*Emotion*] from string
        
        exceptions_pattern =  "?!" + '|'.join(f"{re.escape(exception)}" for exception in self.FILTER_EXEPTIONS)
        pattern = re.compile(r'\[((' + exceptions_pattern + r').*?)\]')
        clean_text =  re.sub(pattern, '', self.text.lower())
        return clean_text
        
    def count_words(self):
        # Remove special characters from text
        return len(self.get_clean_text().split())
    
    def count_chars(self):
        # count letters from A to Z
        letters = re.findall(r'[a-zA-Z0-9]', self.get_clean_text())
        letters_string = ''.join(letters)
        return len(letters)

class AnalyseDocument():
    classifiedParagraphs = []

    def __init__(self, document:Document) -> None:

        self.wordsParent = 0
        self.wordsChild = 0
        self.numberParagraphs = 0
        self.numberParentParagraphs = 0
        self.numberChildParagraphs = 0
        self.lettersParent = 0
        self.lettersChild = 0

        for paragraph in document.paragraphs:
            # categorise sentence
            text = paragraph.text.strip()
            if text != '':
                paragraph = ClassifiedParagraph(text)
                self.classifiedParagraphs.append(paragraph)

                if paragraph.category == Categories.PARENT:
                    self.numberParentParagraphs += 1
                    self.numberParagraphs += 1
                    self.wordsParent += paragraph.numberWords
                    self.lettersParent += paragraph.numberLetters
                
                elif paragraph.category == Categories.CHILDREN:
                    self.numberChildParagraphs += 1
                    self.numberParagraphs += 1
                    self.wordsChild += paragraph.numberWords
                    self.lettersChild += paragraph.numberLetters
    
    def to_csv(self, file):
        with open(file, mode='w', newline='\n') as f:
            header = ['Category', 'Text', 'Words' , 'Letters', 'Open Questions', 'Closed Questions']
            csvWriter = csv.writer(f, delimiter=';', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)

            csvWriter.writerow(header)

            for cp in self.classifiedParagraphs:
                cleaned_text = cp.text.replace('\n', ' ').replace('\r', '')  # Zeilenumbrüche entfernen
                csvWriter.writerow([cp.category.name, cleaned_text, cp.numberWords, cp.numberLetters, None, None])


DOC_FILE_PATH = "Doc Files"
DOC_OUT ="Out"
source = "T1_21060101_ESO_LG.docx"

document = Document(DOC_FILE_PATH + "/" + source) 
analysed = AnalyseDocument(document)
analysed.to_csv(DOC_OUT + "/" + source.replace(".docx", ".csv"))



print("END")