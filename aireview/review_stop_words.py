#!/usr/bin/env python3
"""Review concordance terms and add approved stop words to the local list."""

import argparse
import json
import random
import re
from pathlib import Path

from docx import Document


SCRIPT_DIR = Path(__file__).resolve().parent
STOP_WORD_FILE = SCRIPT_DIR / "stop_word_list.txt"
ENTRY_RE = re.compile(r"^(.+?)\s{2,}(\d+(?:\.\d+)?(?:,\s*\d+(?:\.\d+)?)*)$")


def normalize_term(value):
    return re.sub(r"\s+", " ", value.strip().lower().replace("’", "'"))


def find_concordance_file(path_arg):
    if path_arg:
        path = Path(path_arg).expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"Concordance file not found: {path}")
        return path

    candidates = [
        path
        for path in Path.cwd().glob("concordance_*.docx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            "No concordance_*.docx file found in the current directory. "
            "Pass the generated DOCX path explicitly."
        )
    return max(candidates, key=lambda path: path.stat().st_mtime)


def read_concordance_terms(docx_path):
    document = Document(docx_path)
    terms = []
    seen = set()

    for paragraph in document.paragraphs:
        match = ENTRY_RE.fullmatch(paragraph.text.strip())
        if not match:
            continue
        term = normalize_term(match.group(1))
        if term and term not in seen:
            terms.append(term)
            seen.add(term)

    return terms


def ask_ai(terms):
    """Use aiReview.py's existing Perplexity client; None triggers manual mode."""
    try:
        from aiReview import query_perplexity
    except Exception as error:
        print(f"AI review unavailable ({error}); switching to individual review.")
        return None

    prompt = (
        "Review these terms from a technical manual's concordance. For each exact "
        "term, decide whether it is a generic word that should be omitted from a "
        "technical concordance (decision 'stop') or a meaningful term to retain "
        "(decision 'keep'). Err toward keeping technical terms, acronyms, and "
        "specific concepts. Do not invent or change terms. Return only valid JSON "
        "with this schema: {\"items\":[{\"term\":\"exact input term\","
        "\"decision\":\"stop or keep\",\"reason\":\"brief reason\"}]}\n\n"
        f"Terms: {json.dumps(terms, ensure_ascii=False)}"
    )

    try:
        response = query_perplexity(prompt)
    except Exception as error:
        print(f"AI review failed ({error}); switching to individual review.")
        return None

    return parse_ai_decisions(response, terms)


def parse_ai_decisions(response, terms):
    if not response:
        return None

    text = response.strip()
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        return None

    try:
        payload = json.loads(match.group(0))
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict) or not isinstance(payload.get("items"), list):
        return None

    valid_terms = {normalize_term(term) for term in terms}
    decisions = {}
    for item in payload.get("items", []):
        if not isinstance(item, dict):
            continue
        term = normalize_term(str(item.get("term", "")))
        decision = str(item.get("decision", "")).strip().lower()
        if term not in valid_terms or decision not in {"stop", "keep"}:
            continue
        decisions[term] = {
            "decision": decision,
            "reason": str(item.get("reason", "")).strip(),
        }
    return decisions or None


def individual_review(terms, ai_decisions):
    stop_terms = set()
    keep_terms = set()
    reviewed_terms = set()
    quit_requested = False
    print(
        "\nReview each term: [s]top word, [n]normal/keep, "
        "[k]keep as-is and mark reviewed, [q]uit."
    )

    for term in terms:
        suggestion = ai_decisions.get(term, {})
        if suggestion:
            detail = f"AI: {suggestion['decision']}"
            if suggestion["reason"]:
                detail += f" — {suggestion['reason']}"
        else:
            detail = "AI: no suggestion"

        while True:
            choice = input(f"\n{term} ({detail}) [s/n/k/q]: ").strip().lower()
            if choice in {"s", "n", "k", "q"}:
                break
            print("Choose s, n, k, or q.")

        if choice == "q":
            quit_requested = True
            break
        reviewed_terms.add(term)
        if choice == "s":
            stop_terms.add(term)
        elif choice == "n":
            keep_terms.add(term)

    return stop_terms, keep_terms, reviewed_terms, quit_requested


def write_decisions(stop_terms, keep_terms):
    lines = STOP_WORD_FILE.read_text(encoding="utf-8").splitlines()
    additions = []
    kept_lines = []
    current_terms = set()

    for line in lines:
        term = normalize_term(line.split("#", 1)[0])
        if term and term in keep_terms:
            continue
        kept_lines.append(line)
        if term:
            current_terms.add(term)

    for term in sorted(stop_terms):
        if term not in current_terms:
            additions.append(term)

    if not additions and len(kept_lines) == len(lines):
        print("\nNo stop-word list changes to save.")
        return

    output_lines = kept_lines
    if additions:
        if output_lines and output_lines[-1].strip():
            output_lines.append("")
        output_lines.extend(additions)

    temporary_path = STOP_WORD_FILE.with_name(STOP_WORD_FILE.name + ".tmp")
    temporary_path.write_text("\n".join(output_lines).rstrip() + "\n", encoding="utf-8")
    temporary_path.replace(STOP_WORD_FILE)
    print(
        f"\nSaved {len(additions)} new stop word(s) to {STOP_WORD_FILE.name}."
        + (f" Removed {len(keep_terms)} reviewed normal word(s) from that list." if keep_terms else "")
    )


def main():
    parser = argparse.ArgumentParser(
        description="AI-assisted review of generated concordance terms."
    )
    parser.add_argument(
        "docx",
        nargs="?",
        help="Generated concordance DOCX (defaults to the newest concordance_*.docx).",
    )
    parser.add_argument(
        "--sample-size",
        type=int,
        default=10,
        help="Number of random terms to review (default: 10).",
    )
    parser.add_argument("--seed", type=int, help="Optional random seed for repeatable sampling.")
    parser.add_argument(
        "--no-ai",
        action="store_true",
        help="Skip AI and review the sampled terms individually.",
    )
    args = parser.parse_args()

    if args.sample_size < 1:
        parser.error("--sample-size must be at least 1")

    try:
        if not STOP_WORD_FILE.is_file():
            raise FileNotFoundError(f"Stop-word list not found: {STOP_WORD_FILE}")
        docx_path = find_concordance_file(args.docx)
        terms = read_concordance_terms(docx_path)
    except (FileNotFoundError, OSError, ValueError) as error:
        parser.error(str(error))

    if not terms:
        print("No concordance terms found in the document.")
        return

    rng = random.Random(args.seed)
    print(f"Concordance: {docx_path}")
    reviewed_terms = set()
    batch_number = 0

    while len(reviewed_terms) < len(terms):
        pending_terms = [term for term in terms if term not in reviewed_terms]
        sample = rng.sample(pending_terms, min(args.sample_size, len(pending_terms)))
        batch_number += 1
        print(
            f"\n--- Batch {batch_number}: {len(sample)} terms "
            f"({len(reviewed_terms)}/{len(terms)} already reviewed) ---"
        )
        for term in sample:
            print(f"  - {term}")

        decisions = None if args.no_ai else ask_ai(sample)
        if decisions:
            print("\nAI review:")
            for term in sample:
                suggestion = decisions.get(
                    term,
                    {"decision": "keep", "reason": "No AI decision; defaulting to keep."},
                )
                reason = f" — {suggestion['reason']}" if suggestion["reason"] else ""
                print(f"  {term}: {suggestion['decision']}{reason}")

            while True:
                choice = input(
                    "\n[a]ccept AI decisions, [r]eview each term, [q]uit: "
                ).strip().lower()
                if choice in {"a", "r", "q"}:
                    break
                print("Choose a, r, or q.")

            if choice == "q":
                break
            if choice == "a":
                stop_terms = {
                    term
                    for term in sample
                    if decisions.get(term, {}).get("decision") == "stop"
                }
                keep_terms = set(sample) - stop_terms
                write_decisions(stop_terms, keep_terms)
                reviewed_terms.update(sample)
                continue
        else:
            print("\nNo usable AI response; reviewing the sample individually.")
            decisions = {}

        stop_terms, keep_terms, batch_reviewed, quit_requested = individual_review(
            sample, decisions
        )
        if stop_terms or keep_terms:
            write_decisions(stop_terms, keep_terms)
        reviewed_terms.update(batch_reviewed)
        print(f"Reviewed {len(reviewed_terms)}/{len(terms)} terms so far.")
        if quit_requested:
            break

    if len(reviewed_terms) == len(terms):
        print(f"\nReview complete: all {len(terms)} concordance terms reviewed.")
    else:
        print(
            f"\nReview paused: {len(reviewed_terms)}/{len(terms)} terms reviewed. "
            "Run the command again to start a new review session."
        )


if __name__ == "__main__":
    main()
