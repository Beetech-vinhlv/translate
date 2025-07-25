class Translator:
    def __init__(self, dictionary):
        self.dictionary = dictionary

    def translate(self, text):
        return self.dictionary.get(text, text)