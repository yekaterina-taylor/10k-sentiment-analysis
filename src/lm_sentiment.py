"""Score MD&A text with the Loughran-McDonald (LM) sentiment word lists.

For every <TICKER>_FY<year>_MDA_SentimentReady.txt file in a folder this counts:
  total words, Positive, Negative, Uncertainty, Litigious, Strong modal,
  Weak modal and Constraining words, then derives
  P/N ratio        = Positive / Negative
  Sentiment score  = (Positive - Negative) / Total words

Method (kept identical for every company so results are comparable):
  * Text is lowercased; a "word" is a run of letters, optionally with one
    apostrophe part (e.g. "company's"). Numbers and symbols are not words.
  * A word counts in a category when it appears, exactly, in that LM list.
    No stemming: LM already lists each inflected form separately.
  * Strong modal = LM "Modal" code 1, Weak modal = LM "Modal" code 3.

Dictionary: LM_MasterDictionary.csv (the Loughran-McDonald Master Dictionary
as distributed with the pysentiment2 package, 86,486 words).

Usage: python lm_sentiment.py <TICKER> <mda_folder> [output.json]
"""
import csv
import json
import re
import sys
from pathlib import Path

DICT_FILE = Path(__file__).with_name("LM_MasterDictionary.csv")
TOKEN = re.compile(r"[a-z]+(?:'[a-z]+)?")
CATEGORIES = ["positive", "negative", "uncertainty", "litigious",
              "strongModal", "weakModal", "constraining"]


def load_lists(path=DICT_FILE):
    lists = {c: set() for c in CATEGORIES}
    with open(path, encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            w = row["Word"].lower()
            if row["Positive"] != "0":
                lists["positive"].add(w)
            if row["Negative"] != "0":
                lists["negative"].add(w)
            if row["Uncertainty"] != "0":
                lists["uncertainty"].add(w)
            if row["Litigious"] != "0":
                lists["litigious"].add(w)
            if row["Constraining"] != "0":
                lists["constraining"].add(w)
            if row["Modal"] == "1":
                lists["strongModal"].add(w)
            if row["Modal"] == "3":
                lists["weakModal"].add(w)
    return lists


def score_text(text, lists):
    tokens = TOKEN.findall(text.lower())
    out = {"total": len(tokens)}
    for c in CATEGORIES:
        out[c] = sum(1 for t in tokens if t in lists[c])
    out["pn"] = out["positive"] / out["negative"] if out["negative"] else None
    out["score"] = (out["positive"] - out["negative"]) / out["total"] if out["total"] else None
    return out


def main(ticker, folder, out_file=None):
    lists = load_lists()
    rows = []
    for f in sorted(Path(folder).glob(f"{ticker}_FY*_MDA_SentimentReady.txt"), reverse=True):
        year = int(re.search(r"FY(\d{4})", f.name).group(1))
        rows.append({"year": year, **score_text(f.read_text(encoding="utf-8"), lists)})
    for r in rows:
        print(f"FY{r['year']}  words={r['total']:5d}  +{r['positive']:3d}  -{r['negative']:3d}  "
              f"P/N={r['pn']:.2f}  score={r['score']:+.4f}")
    if out_file:
        Path(out_file).write_text(json.dumps(rows, indent=1))
    return rows


if __name__ == "__main__":
    main(sys.argv[1].upper(), sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
