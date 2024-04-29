import csv
import unicodedata
import re
import os
import json
from docx import Document
from enum import Enum

# Globals
SETTINGS_PATH = "settings.json"
DOC_OUT ="Out"

def get_docs_files(pathToDocsFiles):
    # List to store the filenames of Word documents
    fileNames = []

    # Iterate through all files in the directory
    for filename in os.listdir(pathToDocsFiles):
        if filename.endswith(".docx"):
            # Add the filename to the list
            fileNames.append(filename)
    
    return fileNames

def read_settings():
     # Öffnen Sie die JSON-Datei im Lesemodus
    with open(SETTINGS_PATH, "r") as json_file:
        # Laden Sie die Daten aus der JSON-Datei in ein Python-Dictionary
        data = json.load(json_file)
    
    return data

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



def run():
    print("Start run")
    SETTINGS = read_settings()
    docFileNames = get_docs_files(SETTINGS['docFilePath'])
    print(str(docFileNames))

    for docFileName in docFileNames:
        print("Read " + docFileName)
        document = Document(SETTINGS['docFilePath'] + docFileName) 
        csvFileName = docFileName.replace(".docx", ".csv")
        analysed = AnalyseDocument(document)
        print("Create csv: " + csvFileName)
        analysed.to_csv(SETTINGS['csvFiles'] + csvFileName)
    
    print("End run")

run()

