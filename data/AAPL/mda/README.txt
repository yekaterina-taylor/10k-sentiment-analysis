Apple Inc. (AAPL) - MD&A text files, FY2016-FY2025
===================================================

Each fiscal year has two files:

  AAPL_FY<year>_MDA_Cleaned.txt
      Item 7 (MD&A) text as extracted and cleaned by the original program:
      lowercased, HTML removed. Table cells remain as short number lines.

  AAPL_FY<year>_MDA_SentimentReady.txt
      Same text prepared for dictionary scoring (src/prepare_mda.py):
      table fragments and page headers removed, broken sentences rejoined.
      Wording is unchanged, so Loughran-McDonald word counts are unaffected.

Known limitations of the source extraction
  * FY2016-FY2019: fiscal-year numbers were dropped in 6-15 places
    (e.g. "net sales ... for , and"), and some short narrative paragraphs
    next to tables appear to be missing.
  * Table numbers made up roughly 10-13% of the original "Total words" count.

Source: Apple Inc. Form 10-K filings, SEC EDGAR (CIK 0000320193).
