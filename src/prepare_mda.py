"""Turn the provided 'cleaned' MD&A text files into sentiment-ready text.

What this does (and why):
  * Drops table fragments: lines with no real words (numbers, $, %, parentheses,
    footnote markers). They are not language and inflate the total word count.
  * Drops repeated page headers ("table of contents").
  * Re-joins sentences that the HTML extraction split across several lines.
  * Repairs words that the SEC page split in two (e.g. "refe" / "r to note 4"),
    but only when the joined word is in the Loughran-McDonald word list and at
    least one of the two pieces is not a word on its own.
  * Leaves the wording itself untouched, so dictionary counts are not altered.

Usage: python prepare_mda.py <TICKER> <input_dir> <output_dir>
"""
import json
import re
import sys
from pathlib import Path

DICT_FILE = Path(__file__).with_name("LM_MasterDictionary.csv")


def load_vocab():
    import csv
    with open(DICT_FILE, encoding="utf-8") as fh:
        return frozenset(r["Word"].lower() for r in csv.DictReader(fh))


WORD = re.compile(r"[a-z]+(?:['-][a-z]+)*")
END_PUNCT = (".", ":", ";", "?", "!")


def is_table_fragment(line: str) -> bool:
    """True when the line has no word of two or more letters."""
    return not re.search(r"[a-z]{2,}", line)


def rejoin(lines, vocab=frozenset()):
    """Join lines that are continuations of the previous sentence."""
    out = []
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if out:
            prev = out[-1]
            head = re.search(r"([a-z]+)$", prev)
            tail = re.match(r"[a-z]+", line)
            if head and tail:
                a, b = head.group(1), tail.group(0)
                if (a + b) in vocab and (a not in vocab or b not in vocab):
                    out[-1] = prev + line  # word split across lines
                    continue
            continuation = (
                line[0] in ',.;:)%"\'' or            # starts with punctuation
                prev.endswith(('"', "(", "$", ",")) or  # previous clearly unfinished
                (len(prev) > 60 and not prev.endswith(END_PUNCT))  # long unfinished sentence
            )
            if continuation:
                sep = "" if line[0] in ',.;:)%' or prev.endswith(("(", '"')) and line[0] == '"' else " "
                out[-1] = (prev + sep + line).strip()
                continue
        out.append(line)
    # tidy spacing around stray quotes and punctuation
    cleaned = []
    for line in out:
        line = re.sub(r'\s+"\s*', ' "', line)
        line = re.sub(r'\s+([,.;:)])', r"\1", line)
        line = re.sub(r"\s{2,}", " ", line).strip()
        cleaned.append(line)
    return cleaned


def missing_year_gaps(text: str) -> int:
    """Count spots where fiscal years were dropped, e.g. 'for , and (dollars...'."""
    return len(re.findall(r"\b(?:for|during|of|in)\s*,\s*and\b", re.sub(r"\s+", " ", text)))


def main(ticker: str, src: Path, dst: Path):
    dst.mkdir(parents=True, exist_ok=True)
    report = []
    vocab = load_vocab()
    for f in sorted(src.glob(f"*{ticker}_*_MDA_Cleaned.txt")):
        year = int(re.search(ticker + r"_(?:FY)?(\d{4})", f.name).group(1))
        text = f.read_text(encoding="utf-8")
        lines = text.splitlines()
        kept = [l for l in lines if not is_table_fragment(l) and l.strip() != "table of contents"]
        ready = "\n".join(rejoin(kept, vocab)) + "\n"

        (dst / f"{ticker}_FY{year}_MDA_Cleaned.txt").write_text(text, encoding="utf-8")
        (dst / f"{ticker}_FY{year}_MDA_SentimentReady.txt").write_text(ready, encoding="utf-8")
        report.append({
            "year": year,
            "tokensProvided": len(text.split()),
            "wordsSentimentReady": len(WORD.findall(ready)),
            "tableLinesRemoved": len(lines) - len(kept),
            "missingYearGaps": missing_year_gaps(text),
        })
    (dst / "mda_report.json").write_text(json.dumps(report, indent=1))
    for r in report:
        print(r)


if __name__ == "__main__":
    main(sys.argv[1].upper(), Path(sys.argv[2]), Path(sys.argv[3]))
