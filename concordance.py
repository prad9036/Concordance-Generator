import sys
import re
import json
import html
from collections import defaultdict, Counter

import fitz  # PyMuPDF


# ============================================================
# SETTINGS
# ============================================================

MIN_WORD_LENGTH = 3

STOP_WORDS = {
    "a", "about", "above", "after", "again", "against", "all",
    "am", "an", "and", "any", "are", "as", "at", "be",
    "because", "been", "before", "being", "below", "between",
    "both", "but", "by", "can", "could", "did", "do", "does",
    "doing", "down", "during", "each", "few", "for", "from",
    "further", "had", "has", "have", "having", "he", "her",
    "here", "hers", "herself", "him", "himself", "his", "how",
    "i", "if", "in", "into", "is", "it", "its", "itself",
    "just", "me", "more", "most", "my", "myself", "no", "nor",
    "not", "now", "of", "off", "on", "once", "only", "or",
    "other", "our", "ours", "ourselves", "out", "over", "own",
    "same", "she", "should", "so", "some", "such", "than",
    "that", "the", "their", "theirs", "them", "themselves",
    "then", "there", "these", "they", "this", "those",
    "through", "to", "too", "under", "until", "up", "very",
    "was", "we", "were", "what", "when", "where", "which",
    "while", "who", "whom", "why", "will", "with", "would",
    "you", "your", "yours", "yourself", "yourselves"
}


# ============================================================
# PDF
# ============================================================

def extract_pages(filename):

    doc = fitz.open(filename)

    pages = []

    print("Reading PDF...")

    for number, page in enumerate(doc, 1):

        pages.append({
            "page": number,
            "text": page.get_text("text")
        })

        if number % 25 == 0:
            print(f"  {number} pages processed")

    doc.close()

    print(f"Total pages: {len(pages)}")

    return pages


# ============================================================
# HEADER / FOOTER
# ============================================================

def find_repeated_lines(pages):

    counter = Counter()

    for item in pages:

        lines = item["text"].splitlines()

        candidates = (
            lines[:4] +
            lines[-4:]
        )

        for line in candidates:

            line = re.sub(
                r"\s+",
                " ",
                line
            ).strip().lower()

            if len(line) >= 4:
                counter[line] += 1

    threshold = max(
        3,
        int(len(pages) * 0.10)
    )

    return {
        line
        for line, count in counter.items()
        if count >= threshold
    }


def remove_headers_footers(
    text,
    repeated_lines
):

    result = []

    for line in text.splitlines():

        normalized = re.sub(
            r"\s+",
            " ",
            line
        ).strip().lower()

        if normalized in repeated_lines:
            continue

        result.append(line)

    return "\n".join(result)


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):

    # Join hyphenated words split over lines.

    text = re.sub(
        r"(\w)-\s*\n\s*(\w)",
        r"\1\2",
        text
    )

    # Replace line breaks with spaces.

    text = re.sub(
        r"\s*\n\s*",
        " ",
        text
    )

    # Collapse spaces.

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# SENTENCES
# ============================================================

def split_sentences(text):

    parts = re.split(
        r"(?<=[.!?])\s+(?=[A-Z0-9\"'])",
        text
    )

    return [
        x.strip()
        for x in parts
        if x.strip()
    ]


# ============================================================
# WORDS
# ============================================================

def extract_words(text):

    return re.findall(
        r"\b[A-Za-z]+(?:['’-][A-Za-z]+)*\b",
        text
    )


def normalize_word(word):

    word = word.lower()

    word = word.replace("’", "'")

    if word.endswith("'s"):
        word = word[:-2]

    word = word.strip("-'")

    return word


# ============================================================
# CONCORDANCE
# ============================================================

def build_concordance(pages):

    repeated = find_repeated_lines(pages)

    print(
        f"Possible repeated headers/footers: "
        f"{len(repeated)}"
    )

    concordance = defaultdict(list)

    total = 0

    for index, item in enumerate(pages, 1):

        text = remove_headers_footers(
            item["text"],
            repeated
        )

        text = clean_text(text)

        if not text:
            continue

        sentences = split_sentences(text)

        for sentence in sentences:

            words = extract_words(sentence)

            for raw_word in words:

                word = normalize_word(
                    raw_word
                )

                if len(word) < MIN_WORD_LENGTH:
                    continue

                if word in STOP_WORDS:
                    continue

                concordance[word].append({
                    "page": item["page"],
                    "context": sentence
                })

                total += 1

        if index % 25 == 0:
            print(
                f"  Indexing {index}/{len(pages)}"
            )

    return concordance, total


# ============================================================
# REMOVE DUPLICATES
# ============================================================

def remove_duplicates(concordance):

    for word in concordance:

        seen = set()
        result = []

        for item in concordance[word]:

            key = (
                item["page"],
                item["context"]
            )

            if key not in seen:

                seen.add(key)
                result.append(item)

        concordance[word] = result


# ============================================================
# HTML TEMPLATE
# ============================================================

HTML_TEMPLATE = r"""
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1">

<title>Book Concordance</title>

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    background: #f5f5f5;
    color: #222;
    font-family:
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        Arial,
        sans-serif;
}

header {
    position: sticky;
    top: 0;
    z-index: 10;

    background: white;

    border-bottom: 1px solid #ddd;

    padding: 20px;
}

.header-inner {
    max-width: 1100px;
    margin: auto;
}

h1 {
    margin: 0;
}

.filename {
    color: #777;
    font-size: 14px;
    margin-top: 5px;
}

.stats {
    color: #777;
    font-size: 13px;
    margin-top: 3px;
}

.search-row {
    display: flex;
    gap: 8px;
    margin-top: 18px;
}

#search {
    flex: 1;

    padding: 13px;

    font-size: 17px;

    border: 1px solid #bbb;

    border-radius: 8px;

    outline: none;
}

#search:focus {
    border-color: #333;
}

#clear {
    padding: 0 18px;

    border: 1px solid #bbb;

    background: white;

    border-radius: 8px;

    cursor: pointer;
}

.info {
    display: flex;

    justify-content: space-between;

    margin-top: 8px;

    color: #777;

    font-size: 13px;
}

.alphabet {
    display: flex;

    flex-wrap: wrap;

    gap: 5px;

    margin-top: 15px;
}

.letter {
    border: 0;

    background: #eee;

    border-radius: 5px;

    padding: 5px 9px;

    cursor: pointer;
}

.letter.active {
    background: #222;
    color: white;
}

main {
    max-width: 1100px;

    margin: auto;

    padding: 25px 20px 80px;
}

.word {
    background: white;

    margin-bottom: 12px;

    padding: 18px 20px;

    border-radius: 10px;

    box-shadow:
        0 1px 3px rgba(0,0,0,.08);
}

.word-name {
    font-size: 21px;

    font-weight: 600;
}

.frequency {
    color: #777;

    font-size: 13px;

    margin-top: 3px;
}

.entry {
    margin-top: 12px;

    padding:
        10px 0 10px 15px;

    border-left:
        3px solid #ddd;

    line-height: 1.5;
}

.page-number {
    font-weight: 600;

    color: #555;

    margin-right: 8px;
}

.context {
    color: #444;
}

mark {
    background: #ffe58a;

    padding: 1px 2px;

    border-radius: 2px;
}

#no-results {
    display: none;

    text-align: center;

    background: white;

    padding: 50px;

    border-radius: 10px;

    color: #777;
}

</style>

</head>


<body>


<header>

<div class="header-inner">

<h1>Concordance</h1>

<div class="filename">
__PDF_NAME__
</div>

<div class="stats">
__UNIQUE_WORDS__ unique words ·
__OCCURRENCES__ occurrences
</div>


<div class="search-row">

<input
    id="search"
    type="search"
    placeholder="Search word or context..."
    autocomplete="off"
>

<button id="clear">
Clear
</button>

</div>


<div class="info">

<span id="result-count">
Showing all words
</span>

<span>
Press Esc to clear
</span>

</div>


<div class="alphabet">

__ALPHABET__

</div>

</div>

</header>


<main>

<div id="results"></div>

<div id="no-results">

<h2>No results found</h2>

Try another search term.

</div>

</main>


<script id="concordance-data"
        type="application/json">

__DATA__

</script>


<script>

"use strict";


/* ==========================================================
   LOAD DATA
========================================================== */

const DATA_ELEMENT =
    document.getElementById(
        "concordance-data"
    );

const DATA =
    JSON.parse(
        DATA_ELEMENT.textContent
    );


/* ==========================================================
   ELEMENTS
========================================================== */

const searchBox =
    document.getElementById("search");

const clearButton =
    document.getElementById("clear");

const results =
    document.getElementById("results");

const noResults =
    document.getElementById("no-results");

const resultCount =
    document.getElementById("result-count");


/* ==========================================================
   STATE
========================================================== */

let selectedLetter = null;


/* ==========================================================
   ESCAPE HTML
========================================================== */

function escapeHTML(text) {

    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


/* ==========================================================
   ESCAPE REGEX
========================================================== */

function escapeRegex(text) {

    return text.replace(
        /[.*+?^${}()|[\]\\]/g,
        "\\$&"
    );
}


/* ==========================================================
   HIGHLIGHT
========================================================== */

function highlight(text, query) {

    const safeText =
        escapeHTML(text);

    if (!query) {
        return safeText;
    }

    const regex =
        new RegExp(
            "(" +
            escapeRegex(query) +
            ")",
            "gi"
        );

    return safeText.replace(
        regex,
        "<mark>$1</mark>"
    );
}


/* ==========================================================
   GET RESULTS
========================================================== */

function getResults(query) {

    let result = DATA;


    /* Alphabet filter */

    if (selectedLetter) {

        result = result.filter(
            function(item) {

                return (
                    item.word
                        .charAt(0)
                        .toUpperCase()
                    ===
                    selectedLetter
                );

            }
        );
    }


    /* Search */

    if (query) {

        const q =
            query.toLowerCase();

        result = result.filter(
            function(item) {

                /* Search word */

                if (
                    item.word
                        .toLowerCase()
                        .includes(q)
                ) {
                    return true;
                }


                /* Search context */

                return item.occurrences.some(
                    function(occurrence) {

                        return (
                            occurrence.context
                                .toLowerCase()
                                .includes(q)
                        );

                    }
                );

            }
        );
    }


    return result;
}


/* ==========================================================
   RENDER
========================================================== */

function render() {

    const query =
        searchBox.value.trim();

    const data =
        getResults(query);


    results.innerHTML = "";


    /* Result count */

    if (!query && !selectedLetter) {

        resultCount.textContent =
            "Showing all " +
            DATA.length.toLocaleString() +
            " words";

    } else {

        resultCount.textContent =
            data.length.toLocaleString() +
            " matching words";
    }


    /* No results */

    if (data.length === 0) {

        noResults.style.display =
            "block";

        return;

    }


    noResults.style.display =
        "none";


    /* Create results */

    let output = "";


    data.forEach(
        function(item) {

            output +=
                '<section class="word">';


            output +=
                '<div class="word-name">' +
                highlight(
                    item.word,
                    query
                ) +
                '</div>';


            output +=
                '<div class="frequency">' +
                item.count.toLocaleString() +
                (
                    item.count === 1
                    ? " occurrence"
                    : " occurrences"
                ) +
                '</div>';


            item.occurrences.forEach(
                function(occurrence) {

                    output +=
                        '<div class="entry">';

                    output +=
                        '<span class="page-number">' +
                        'Page ' +
                        occurrence.page +
                        '</span>';

                    output +=
                        '<span class="context">' +
                        highlight(
                            occurrence.context,
                            query
                        ) +
                        '</span>';

                    output +=
                        '</div>';
                }
            );


            output +=
                '</section>';

        }
    );


    results.innerHTML =
        output;
}


/* ==========================================================
   SEARCH
========================================================== */

searchBox.addEventListener(
    "input",
    function() {

        render();

    }
);


/* ==========================================================
   CLEAR
========================================================== */

clearButton.addEventListener(
    "click",
    function() {

        searchBox.value = "";

        selectedLetter = null;


        document
            .querySelectorAll(".letter")
            .forEach(
                function(button) {

                    button.classList.remove(
                        "active"
                    );

                }
            );


        render();

        searchBox.focus();

    }
);


/* ==========================================================
   ESC
========================================================== */

document.addEventListener(
    "keydown",
    function(event) {

        if (event.key === "Escape") {

            searchBox.value = "";

            selectedLetter = null;


            document
                .querySelectorAll(".letter")
                .forEach(
                    function(button) {

                        button.classList.remove(
                            "active"
                        );

                    }
                );


            render();

            searchBox.focus();

        }

    }
);


/* ==========================================================
   ALPHABET
========================================================== */

document
    .querySelectorAll(".letter")
    .forEach(
        function(button) {

            button.addEventListener(
                "click",
                function() {

                    const letter =
                        this.dataset.letter;


                    if (
                        selectedLetter === letter
                    ) {

                        selectedLetter =
                            null;

                        this.classList.remove(
                            "active"
                        );

                    } else {

                        selectedLetter =
                            letter;


                        document
                            .querySelectorAll(
                                ".letter"
                            )
                            .forEach(
                                function(b) {

                                    b.classList.remove(
                                        "active"
                                    );

                                }
                            );


                        this.classList.add(
                            "active"
                        );
                    }


                    render();

                }
            );

        }
    );


/* ==========================================================
   INITIAL DISPLAY
========================================================== */

render();

</script>


</body>

</html>
"""


# ============================================================
# CREATE HTML
# ============================================================

def create_html(
    concordance,
    output_file,
    pdf_name
):

    words = sorted(
        concordance.keys()
    )


    data = []

    for word in words:

        data.append({
            "word": word,
            "count": len(
                concordance[word]
            ),
            "occurrences":
                concordance[word]
        })


    json_data = json.dumps(
        data,
        ensure_ascii=False
    )


    # Alphabet buttons

    alphabet = ""

    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":

        alphabet += (
            '<button class="letter" '
            'data-letter="' +
            letter +
            '">' +
            letter +
            '</button>'
        )


    occurrence_count = sum(
        item["count"]
        for item in data
    )


    # Replace placeholders

    document = HTML_TEMPLATE

    document = document.replace(
        "__PDF_NAME__",
        html.escape(pdf_name)
    )

    document = document.replace(
        "__UNIQUE_WORDS__",
        f"{len(words):,}"
    )

    document = document.replace(
        "__OCCURRENCES__",
        f"{occurrence_count:,}"
    )

    document = document.replace(
        "__ALPHABET__",
        alphabet
    )

    document = document.replace(
        "__DATA__",
        json_data
    )


    with open(
        output_file,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(document)


# ============================================================
# MAIN
# ============================================================

def main():

    if len(sys.argv) != 3:

        print()
        print(
            'Usage: python concordance.py '
            '"book.pdf" "concordance.html"'
        )
        print()

        sys.exit(1)


    pdf_file = sys.argv[1]

    output_file = sys.argv[2]


    print()
    print("=" * 50)
    print("PDF CONCORDANCE")
    print("=" * 50)
    print()


    # Read PDF

    pages = extract_pages(
        pdf_file
    )


    # Build concordance

    concordance, total = \
        build_concordance(pages)


    # Remove duplicate contexts

    print("Removing duplicates...")

    remove_duplicates(
        concordance
    )


    # Generate HTML

    print("Creating HTML...")

    create_html(
        concordance,
        output_file,
        pdf_file
    )


    print()
    print("DONE")
    print()
    print(
        f"Unique words: "
        f"{len(concordance):,}"
    )

    print(
        f"Indexed occurrences: "
        f"{total:,}"
    )

    print(
        f"Output: {output_file}"
    )

    print()


if __name__ == "__main__":
    main()
