# Installation

Download the requirements from `requirements.txt` using the command

```bash
pip install -r ./requirements.txt
```

After installation install the wordlists of each of the libraries. To do so use commands:

```bash
python -m nltk.downloader wordnet omw-1.4
python -m spacy download en_core_web_sm
```

# File format

**Note: Table headers are to be present in the XLSXs**
`word_frequencies.xlsx` must have the following format
| Word | Frequency |
|-------|----------|
| word1 | 1 |
| word2 | 2 |

`GRE Word sheet.xlsx` only have the first columnt to match the below table, the rest doesn't matter as only the first row will be used for reordering/bucketing
| Word | Column1 | Column 2 |
|-------|----------|---|
| word1 | value1 | value11 |
| word2 | value2 | value21 |
