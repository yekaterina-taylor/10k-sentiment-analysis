NVIDIA Corp. (NVDA) - MD&A text files, FY2017-FY2026
=====================================================

Each fiscal year has three files:

  NVDA_FY<year>_MDA_Raw.txt
      Item 7 (MD&A) text as extracted from the 10-K: original capitalization,
      page numbers and "Table of Contents" headers, one table cell per line.

  NVDA_FY<year>_MDA_Cleaned.txt
      The raw text lowercased, with page numbers removed and quotes normalized.
      Table cells remain as short lines of numbers.

  NVDA_FY<year>_MDA_SentimentReady.txt
      The cleaned text prepared for dictionary scoring (src/prepare_mda.py):
      table fragments and page headers removed, broken sentences rejoined, and
      words the SEC page split in two repaired (e.g. "refe" + "r" -> "refer").
      Wording is otherwise unchanged.

Checks
  * All ten files are complete: each starts at the Item 7 heading and ends with
    "Adoption of New and Recently Issued Accounting Pronouncements".
  * The cleaning step dropped standalone years and small numbers
    (e.g. "fiscal year 2017" became "fiscal year"). Numbers are not counted as
    words, so this does not change sentiment counts.
  * FY2017-FY2019 MD&As compare three years; from FY2020 NVIDIA discusses only
    the latest two years and refers back to the prior 10-K, so the text is shorter.

Source: NVIDIA Corp. Form 10-K filings, SEC EDGAR (CIK 0001045810).
