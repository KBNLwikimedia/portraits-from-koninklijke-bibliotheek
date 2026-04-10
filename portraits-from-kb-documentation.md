# Portraits from Koninklijke Bibliotheek — Documentation

## Overview

This project analyses and enriches structured data on Wikimedia Commons files in the category [Portraits from Koninklijke Bibliotheek](https://commons.wikimedia.org/wiki/Category:Portraits_from_Koninklijke_Bibliotheek). It checks whether each file has a [depicts (P180)](https://www.wikidata.org/wiki/Property:P180) statement pointing to a human ([Q5](https://www.wikidata.org/wiki/Q5)), and collects metadata such as file captions, categories, and Wikidata usage.

## Excel workbook

**File:** `kb-portraits-depicts-q5-status-YYYYMMDD.xlsx`

The workbook contains three sheets, all sharing the same 12-column structure.

### Sheets

| Sheet | Description | Rows (10 Apr 2025) |
|---|---|---|
| `all` | All files in the category — one row per depicted human per file. Files without a depicts statement have empty depicts columns. | 3,223 |
| `missing` | Subset of `all`: files that have **no** depicts (P180) statement pointing to a human. | 95 |
| `present` | Subset of `all`: files that **have** at least one depicts (P180) statement pointing to a human. | 3,128 |

### Columns

| # | Column | Description |
|---|---|---|
| 1 | `file` | Commons entity URL, e.g. `https://commons.wikimedia.org/entity/M12345` |
| 2 | `title` | Full file title on Commons, e.g. `File:Portrait of Example.jpg` |
| 3 | `file_caption_en` | English file caption (MediaInfo label) on Commons |
| 4 | `file_caption_nl` | Dutch file caption (MediaInfo label) on Commons |
| 5 | `depicts_P180` | Wikidata Q-id of the depicted human, e.g. `Q12345` |
| 6 | `depicts_P180_label_en` | English label of the depicted person on Wikidata |
| 7 | `depicts_P180_description_en` | English description of the depicted person on Wikidata |
| 8 | `depicts_P180_label_nl` | Dutch label of the depicted person on Wikidata |
| 9 | `depicts_P180_description_nl` | Dutch description of the depicted person on Wikidata |
| 10 | `categories` | Non-hidden Commons categories (semicolon-separated) |
| 11 | `wikidata_usage` | Wikidata Q-ids where this file is used (semicolon-separated) |
| 12 | `wikidata_label` | English labels for the Wikidata items in `wikidata_usage` (semicolon-separated) |

**Note:** Files with multiple depicts (P180) values pointing to different humans are split into multiple rows (one per depicted person). All other columns are repeated for each row.

## Python script

**File:** `portraits-from-kb-assess-depicts-q5-status.py`

### What it does

The script fetches all files from the Commons category, extracts their structured data, and writes the results to CSV and Excel files.

### Workflow

| Step | Description |
|---|---|
| 0 | **Authenticate** with Wikimedia Commons using bot credentials from `.env` |
| 1 | **Fetch category members** — retrieves all files from "Portraits from Koninklijke Bibliotheek" via the MediaWiki API, handling pagination automatically |
| 2 | **Check P180 depicts** — calls `wbgetentities` on Commons in batches to extract depicts (P180) Q-ids and EN/NL file captions from each file's MediaInfo entity |
| 2b | **Filter to humans** — queries Wikidata for all depicts Q-ids in one pass to check P31=Q5 (instance of human) and retrieve EN + NL labels and descriptions. Non-human depicts items are discarded |
| 3 | **Split files** into with/without depicts |
| 4 | **Fetch file details** — retrieves non-hidden categories and global usage (links to Wikidata items) for all files |
| 5 | **Fetch Wikidata labels** — retrieves English labels for Q-ids found in global usage |
| 6 | **Build result rows** — one row per depicted human per file, merging all collected data |
| 7 | **Write CSV files** — `kb-portraits-depicts-q5-status-all-YYYYMMDD.csv` and `kb-portraits-depicts-q5-status-missing-YYYYMMDD.csv` |
| 8 | **Write Excel workbook** — `kb-portraits-depicts-q5-status-YYYYMMDD.xlsx` with three sheets: `all`, `missing`, `present` |

### API details

- **Commons API** (`https://commons.wikimedia.org/w/api.php`): Used for authentication, fetching category members, retrieving MediaInfo entities (structured data + file captions), categories, and global usage.
- **Wikidata API** (`https://www.wikidata.org/w/api.php`): Used for checking instance-of-human (P31=Q5) and fetching labels/descriptions in English and Dutch.
- Commons MediaInfo entities use `statements` (not `claims`) for structured data, and `labels` for file captions.

### Configuration

The script reads credentials from a `.env` file in the project root:

```
COMMONS_USERNAME=YourBotUsername
COMMONS_PASSWORD=YourBotPassword
COMMONS_USER_AGENT=YourBot/1.0
```

### Dependencies

- Python 3.7+
- `requests`
- `python-dotenv`
- `openpyxl`

## SPARQL queries

**File:** `kb-portraits-depicts-status-wmc-queries.rq`

This file contains two SPARQL queries that can be run on the [Wikimedia Commons Query Service](https://commons-query.wikimedia.org/). Both queries retrieve files from the "Portraits from Koninklijke Bibliotheek" category and join them with Wikidata to find depicted humans. They use a federated query pattern: first fetching files from Commons via `wikibase:mwapi`, then crossing to Wikidata via `SERVICE <https://query.wikidata.org/sparql>` to filter for humans (P31=Q5).

### Query 1 — Grouped overview

Returns one row per file with all depicted humans and their Commons categories concatenated into pipe-separated lists.

| Column | Description |
|---|---|
| `file` | Commons entity URI (e.g. `https://commons.wikimedia.org/entity/M12345`) |
| `title` | File title on Commons |
| `depicts_list` | Pipe-separated Wikidata URIs of depicted humans |
| `depicts_labels` | Pipe-separated English labels of depicted humans |
| `commonscats` | Pipe-separated Commons categories (P373) of depicted humans |

Key features:
- Uses `GROUP_CONCAT` to aggregate multiple depicted persons per file into a single row.
- Only includes depicted items that are instances of human (Q5) **and** have a Commons category (P373).

### Query 2 — Flat listing

Returns one row per depicted human per file (i.e. files with multiple depicts values produce multiple rows). This is the same structure as the Excel output from the Python script.

| Column | Description |
|---|---|
| `file` | Commons entity URI |
| `title` | File title on Commons |
| `depicts` | Wikidata URI of the depicted human |
| `depicts_label` | English label of the depicted human |
| `depicts_description` | English description of the depicted human |
| `commonscat` | Commons category (P373) of the depicted human, if available |

Key features:
- Only returns files that **have** a depicts (P180) statement pointing to a human (no rows for files without depicts).
- Commons category, label, and description are all `OPTIONAL`, so rows are returned even when these are missing on Wikidata.