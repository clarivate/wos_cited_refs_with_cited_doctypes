# WoS SR + Cited References + Document Type Analysis

This repository is a variation of the original **WoS SR + Cited References Analysis** project:

https://github.com/clarivate/wose_sr_cr_analysis

Like the original project, this script exports **Web of Science (Expanded API)** *Short Records* (SR) for a query, collects the **UT/UIDs**, retrieves **cited references** for each record, writes a CSV, and produces an Excel workbook with summary analysis.

This variation adds an additional enrichment step: when a cited reference has a valid UID/UT, the script looks up that cited record in **Web of Science All Databases (`WOK`)** using Short Record retrieval and adds the cited record's **document type(s)** to the output.

It is intended to be used as a **script within a package/repo** (not as a standalone pip package), alongside the included robust API helper modules.

## What’s included

- `wos_cited_refs_with_cited_doctypes_batched.py` — top-level runner that:
  - runs a WoS SR search (`optionView=SR`) to collect Source UIDs/UTs
  - calls the cited references endpoint for each Source UID
  - writes a CSV of Source UID → Cited Reference fields
  - identifies cited-reference UIDs that can be looked up
  - retrieves cited-reference Short Records from **WOK / All Databases** in batches
  - adds the cited record's document type(s) to the CSV
  - writes an Excel workbook with summary analysis and detailed cited-reference data
- `wosesrclient_robust.py` — robust SR client with retries/backoff and friendly invalid-query error handling
- `wosereferencesclient_robust.py` — robust cited-references client with retries/backoff
- `JCR 2025.csv` — local mapping file used for “Cited Work” consolidation

> License: MIT (see `LICENSE`)

## Requirements

- Python 3.10+ recommended
- A valid **Clarivate Web of Science Expanded API** key
- Access sufficient to search **WOS** and **WOK / All Databases**

Python dependencies are listed in `requirements.txt`.

## Setup

1. Clone the repository and create a virtual environment.

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the repo root (or set the environment variable in your shell):

```bash
EXPANDED_APIKEY=YOUR_WOS_EXPANDED_API_KEY
```

## Running the script

Basic usage:

```bash
python wos_cited_refs_with_cited_doctypes_batched.py -q "AU=Stanwood"
```

### Options

- `--include-zero-ref-uids`  
  Appends one blank row per Source UID that had **zero** cited references to the output CSV (default: off).

- `--no-excel`  
  Skips Excel output (default: Excel workbook is created).

- `-k` / `--key`  
  Allows a Web of Science Expanded API key to be supplied directly instead of using `EXPANDED_APIKEY` from `.env`.

Example:

```bash
python wos_cited_refs_with_cited_doctypes_batched.py -q "TS=CRISPR" --include-zero-ref-uids
```

### Parameters section (in-script defaults)

At the top of the script you can also set defaults without CLI flags:

```python
INCLUDE_ZERO_REF_UIDS = False
MAKE_EXCEL = True
ADD_CITED_REF_DOCTYPES = True
CITED_REF_LOOKUP_DATABASE = "WOK"
CITED_REF_LOOKUP_BATCH_SIZE = 50
```

`WOK` is intentionally used for the cited-reference document-type lookup because some cited-reference UIDs may be available in **Web of Science All Databases** even when they are not present in the Web of Science Core Collection.

## Cited-reference document type enrichment

After the cited-reference CSV has been created, the script performs an additional Short Record lookup.

### Eligible cited-reference UIDs

Only nonblank `Cited Reference UID` values containing a colon (`:`) are treated as valid UIDs for this enrichment step.

For example:

```text
WOS:001391684300001
```

is eligible for lookup, while a cited-reference identifier without a colon is skipped.

Skipped identifiers remain in the output, but their `Doc Types of Cited Ref` value is left blank.

### Deduplication and batching

The script first deduplicates eligible cited-reference UIDs so that the same cited record is not retrieved repeatedly.

Eligible UIDs are then grouped into batches of up to **50** and searched using WOK Short Records.

Conceptually, a batch query looks like:

```text
UT=(WOS:001391684300001 OR WOS:001234567800001 OR ...)
```

This significantly reduces the number of API calls when the same cited references occur multiple times or when many cited UIDs can be retrieved together.

### Multiple document types

A Web of Science record may have more than one document type.

When multiple document types are returned, they are combined in the output using a semicolon separator, for example:

```text
Article; Proceedings Paper
```

If an eligible UID is not returned from WOK, or no document type is available, the document-type field is left blank.

## Outputs

### CSV

A CSV is written to the working directory, named like:

```text
WOS_CitedRefs_<query>_<timestamp>.csv
```

Columns include:

- Source UID
- Cited Reference UID
- **Doc Types of Cited Ref**
- Cited Author
- Year
- Volume
- Page
- Cited Work
- Cited Title
- DOI

Example:

| Source UID | Cited Reference UID | Doc Types of Cited Ref |
|---|---|---|
| WOS:001769767400001 | WOS:001391684300001 | Article |

### Excel

If enabled (default), an Excel file is written:

```text
CR_Pivot_Analysis_<timestamp>.xlsx
```

Sheets include:

- `Cited Work analysis`
  - cited-work counts
  - percent of cited references
  - cumulative counts and percentages
  - Bradford zones

- `Cited Year analysis`
  - cited-year counts
  - percent of cited references
  - “Top X%” coverage summary

- `Cited Reference Details`
  - full row-level cited-reference data
  - includes `Doc Types of Cited Ref`
  - opens as the active worksheet when the workbook is opened
  - includes Excel filter dropdowns across the header row
  - freezes the header row for easier browsing

- `Search Summary`
  - original query
  - local run timestamp
  - total source records retrieved
  - total cited references retrieved
  - number of source records with zero cited references
  - any cited-work variant merges applied through `JCR 2025.csv`

## JCR mapping file (`JCR 2025.csv`)

The script attempts to load a file named **`JCR 2025.csv`** from the same folder as the script. It expects two columns:

- Column A: variant name
- Column B: canonical name

This mapping is used to merge “Cited Work” variants into canonical titles for analysis.

If the file is missing, the script runs normally; cited-work names are simply not consolidated using the mapping file.

## Console reporting

During the cited-reference document-type enrichment step, the script reports:

- number of distinct cited-reference UIDs eligible for lookup
- number of cited-reference identifiers skipped because they do not contain a colon
- number of WOK batches to retrieve
- number of requested UIDs returned for each batch
- number of returned records with document types
- final counts for:
  - UIDs with document types
  - UIDs with no match/document type
  - lookup errors

These messages are intended to make it easier to monitor larger runs and identify incomplete WOK lookups.

## Notes and limitations

- WoS API limits and rate limits apply.
- The initial source-record search uses **WOS / Web of Science Core Collection**.
- The cited-reference document-type enrichment uses **WOK / Web of Science All Databases**.
- Only cited-reference UIDs containing a colon are submitted for document-type lookup.
- Cited-reference UIDs are deduplicated before lookup.
- WOK Short Record lookups are performed in batches of up to 50 UIDs.
- If a WOK lookup does not return a requested UID, its document type is left blank.
- If a batch lookup fails, the script continues and leaves document types blank for that affected batch.
- The SR client surfaces invalid field-tag searches as a friendly `InvalidWoSQueryError`.
- For very large source queries, consider tightening the query or slicing by year or other fields.

## Relationship to the original project

This project is derived from and extends the original **WoS SR + Cited References Analysis**:

https://github.com/clarivate/wose_sr_cr_analysis

The original workflow retrieves source records, cited references, and the cited-work/cited-year analysis.

This variation preserves that workflow and adds **cited-reference document type enrichment using WOK Short Records**, along with changes to the detailed Excel output to make document-type filtering easier.

## MIT License

This project is released under the MIT License.
