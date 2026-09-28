# PDF Concordance Generator

`concordance.py` extracts text from a PDF book, builds an alphabetical index of the words it finds, and saves the result as a three-column Word document (`.docx`). Each entry lists the PDF page labels where the word appears.

The tool removes frequently repeated header and footer lines, joins words split across line breaks, excludes configured stop words, and uses NLTK to normalize common word forms. It is intended for text-based PDFs; scanned pages without a text layer need OCR before they can be indexed.

## Requirements

- Python 3.9 or newer
- The Python packages in `requirements.txt`
- Internet access on the first run so NLTK can download its language data, unless those resources are already installed
- A PDF with extractable text

## Setup

From the repository directory, create and activate a virtual environment, then install the dependencies:

```bash
cd concordance
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

On Windows PowerShell, activate the environment with:

```powershell
.venv\Scripts\Activate.ps1
```

The script checks for the NLTK `wordnet` and `averaged_perceptron_tagger_eng` resources at runtime. If needed, it downloads them into an `nltk_data` directory under the current working directory. Run it from a directory where you have write permission.

## Usage

Run from the `concordance` directory or provide paths to the script and PDF:

```bash
python concordance.py /path/to/book.pdf
```

By default, the output is written to the current working directory as `concordance_<pdf-filename>.docx`. For example, `books/manual.pdf` produces `concordance_manual.docx` in the directory where the command was run.

To choose an output path:

```bash
python concordance.py /path/to/book.pdf /path/to/output.docx
```

Quote paths containing spaces:

```bash
python concordance.py "books/Field Manual.pdf" "output/Field Manual Index.docx"
```

Open the resulting `.docx` in Microsoft Word, LibreOffice Writer, or another compatible word processor.

## Stop word list

`stop_word_list.txt` contains one word per line. Words on this list are omitted from the concordance. Lines beginning with `#` are comments, and text after `#` on a line is ignored. Edit the file to tune the index for a particular book; the list is loaded from the same directory as `concordance.py`.

The script also omits words shorter than three letters, lowercases entries, removes possessive endings, normalizes some spelling variants, and lemmatizes common inflections. Page labels are inferred from page margins where possible; if no suitable printed label is found, the PDF's page number is used.

## Output format

The generated Word document has a centered title and source filename, an alphabetical list grouped by initial letter, and three page columns. Each bold word is followed by a comma-separated list of page labels. Page layout and font settings are defined near the top of `concordance.py`.

## Notes and limitations

- PDF text extraction quality depends on the PDF. Columns, tables, unusual fonts, and complex layouts can affect word order or page-number detection.
- Repeated header/footer detection is heuristic and may remove a short line that repeats in the first or last four extracted lines of many pages.
- WordNet lemmatization and POS tagging are useful approximations; review the generated index for technical or specialized terminology.
- The script does not perform OCR, preserve the source document's layout, or include word counts or surrounding context.
