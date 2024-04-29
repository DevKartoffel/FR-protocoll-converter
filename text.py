import re

def extract_letters(text):
    # Finden Sie alle Buchstaben im Text
    letters = re.findall(r'[a-zA-Z]', text)
    
    # Fügen Sie die Buchstaben zu einem einzelnen String zusammen
    letters_string = ''.join(letters)
    
    return letters_string

# Beispieltext
text = "Hello, this is a test string with special characters: !@#$%^&*()_+"

# Extrahieren Sie nur Buchstaben aus dem gesamten Satz
letters_only = extract_letters(text)
print("Nur Buchstaben im Text:", letters_only)
