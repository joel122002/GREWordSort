import nltk
import spacy
from nltk.stem import PorterStemmer, WordNetLemmatizer
from nltk.corpus import wordnet as wn
from openpyxl import load_workbook, Workbook
from copy import copy
from openpyxl.cell.read_only import EmptyCell

# Initialize NLP tools
stemmer = PorterStemmer()
lemmatizer = WordNetLemmatizer()
nlp = spacy.load("en_core_web_sm")

def are_derivationally_related(word1, word2):
    synsets1 = wn.synsets(word1)
    related_forms = set()
    for s in synsets1:
        for lemma in s.lemmas():
            related_forms.update([d.name() for d in lemma.derivationally_related_forms()])
    return word2 in related_forms

def have_same_root(word1, word2):
    word1 = word1.lower()
    word2 = word2.lower()

    # 1. Stemming
    if stemmer.stem(word1) == stemmer.stem(word2):
        return True

    # 2. Lemmatization (NLTK)
    if lemmatizer.lemmatize(word1) == lemmatizer.lemmatize(word2):
        return True

    # 3. Lemmatization (spaCy)
    doc = nlp(f"{word1} {word2}")
    lemmas = [token.lemma_ for token in doc]
    if lemmas[0] == lemmas[1]:
        return True

    # 4. WordNet derivationally related
    if are_derivationally_related(word1, word2) or are_derivationally_related(word2, word1):
        return True

    return False

MY_WORDSHEET_PATH = "GRE Word sheet.xlsx"
NEW_WORDSHEET_PATH = "GRE Word sheet - New.xlsx"
HIGH_FREQUENCY_WORDS = ["aesthetic",
"alacrity",
"archaic",
"ascetic",
"assuage",
"audacious",
"austere",
"banal",
"capricious",
"censure",
"coalesce",
"craven",
"demur",
"deride",
"derivative",
"diatribe",
"didactic",
"diffident",
"disparate",
"ephemeral",
"eschew",
"esoteric",
"facetious",
"fortuitous",
"garrulous",
"hackneyed",
"immutable",
"inimical",
"innocuous",
"insipid",
"irascible",
"laconic",
"lucid",
"malleable",
"mercurial",
"meticulous",
"mitigate",
"obsequious",
"obstinate",
"opaque",
"perfunctory",
"phlegmatic",
"platitude",
"pristine",
"prodigal",
"recondite",
"refute",
"repudiate",
"reticent",
"sedulous",
"soporific",
"taciturn"]

def row_contains_highly_frequent_word(row):
    for word in HIGH_FREQUENCY_WORDS:
        try:
            if have_same_root(word, row[0].value):
                return True
        except:
            pass 
    return False

# Load the Excel file
wb = load_workbook(MY_WORDSHEET_PATH, read_only=True)
ws = wb.active  # You can also use wb['SheetName'] if needed

list1, list2 = [], []

for row in ws.iter_rows(min_row=2):
    if row_contains_highly_frequent_word(row):
        list1.append(row)
    else:
        list2.append(row)

# Create output workbook
out_wb = Workbook()
out_ws = out_wb.active
header_row = list(ws.iter_rows(min_row=1, max_row=1))[0]
# Write header with formatting
for col_idx, cell in enumerate(header_row, start=1):
    new_cell = out_ws.cell(row=1, column=col_idx, value=cell.value)
    if not isinstance(cell, EmptyCell) and cell.has_style:
        new_cell.font = copy(cell.font)
        new_cell.fill = copy(cell.fill)
        new_cell.border = copy(cell.border)
        new_cell.alignment = copy(cell.alignment)
        new_cell.number_format = cell.number_format

# Function to copy a row with formatting
def write_row(out_ws, row_index, source_row):
    for col_idx, cell in enumerate(source_row, start=1):
        new_cell = out_ws.cell(row=row_index, column=col_idx, value=cell.value)
        if not isinstance(cell, EmptyCell) and cell.has_style:
            new_cell.font = copy(cell.font)
            new_cell.fill = copy(cell.fill)
            new_cell.border = copy(cell.border)
            new_cell.alignment = copy(cell.alignment)
            new_cell.number_format = cell.number_format

# Write list1 then list2 rows
row_index = 2
for row in list1 + list2:
    write_row(out_ws, row_index, row)
    row_index += 1

# Save the output
out_wb.save(NEW_WORDSHEET_PATH)
print(f"Saved output to '{NEW_WORDSHEET_PATH}'")
