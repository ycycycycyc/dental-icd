# K00-K14 ICD-10 Chinese-English Master Table

This repository builds a structured master table for ICD-10 `K00-K14` diagnoses from the Chinese `ICD-10医保2.0版` source, then adds WHO ICD-10 English mappings, semantic English diagnosis names, and QA review flags.

The current recommended working outputs are:

- `K00-K14_master_table_v3_semantic_en.tsv`
- `K00-K14_master_table_v3_semantic_en.xlsx`
- `K00-K14_master_table_v4_QA.tsv`
- `K00-K14_master_table_v4_QA.xlsx`

## Side Project

- [K00-K14 Dental ICD Lookup / 口腔 ICD-10 双语诊断编码查询](projects/dental-icd-lookup/)

This static GitHub Pages-ready lookup page lets users search the K00-K14 semantic table by China diagnosis code, Chinese term, English semantic diagnosis term, WHO ICD code, and WHO English parent term.

## What To Upload To GitHub

### Recommended First Upload

Upload these first. They are enough for someone to understand and use the current work.

| File | Purpose |
|---|---|
| `README.md` | Project overview and file guide |
| `K00-K14_master_table_v3_semantic_en.tsv` | Frozen semantic English master table |
| `K00-K14_master_table_v3_semantic_en.xlsx` | Excel version of v3 |
| `K00-K14_master_table_v4_QA.tsv` | QA review layer |
| `K00-K14_master_table_v4_QA.xlsx` | Excel version of QA layer |
| `K00-K14_v3_semantic_master_summary_and_data_dictionary.md` | Summary and v3 column dictionary |
| `K00-K14_v4_QA_review_instructions.md` | Manual QA review instructions |

### Recommended Script Upload

Upload these with the first batch if you want the work to be reproducible.

| File | Purpose |
|---|---|
| `extract_k00_k14.py` | Extracts the raw K00-K14 hierarchy from the PDF |
| `build_master_table_v1.py` | Builds cleaned Chinese hierarchy master table |
| `rename_master_table_columns.py` | Renames Chinese headers to English snake_case |
| `build_master_table_v2_who_english.py` | Adds WHO ICD-10 English parent mappings |
| `build_master_table_v3_semantic_en.py` | Adds semantic English diagnosis layer |
| `build_master_table_v4_QA.py` | Adds QA status/reason/priority layer |

### Optional Audit Trail Upload

Upload these if you want reviewers to inspect every intermediate layer.

| File | Purpose |
|---|---|
| `k00_k14_full_hierarchy.tsv` | Raw official hierarchy extraction |
| `k00_k14_full_hierarchy.xlsx` | Excel version of raw hierarchy extraction |
| `K00-K14_master_table_v1.tsv` | Cleaned Chinese hierarchy master table |
| `K00-K14_master_table_v1.xlsx` | Excel version of v1 |
| `K00-K14_master_table_v1_english_columns.tsv` | v1 with English snake_case headers |
| `K00-K14_master_table_v1_english_columns.xlsx` | Excel version of English-header v1 |
| `K00-K14_master_table_v2.tsv` | WHO English parent mapping layer |
| `K00-K14_master_table_v2.xlsx` | Excel version of v2 |

### Do Not Upload By Default

These should usually stay out of GitHub.

| File | Reason |
|---|---|
| `ICD-10医保2.0版.pdf` | Original source PDF; upload only if redistribution is permitted |
| `k00_k14.tsv` | Old two-column extraction, superseded by full hierarchy files |
| `__pycache__/` | Python cache |
| `~$*.xlsx` | Temporary Excel lock files |

## Current Data Layers

### Raw hierarchy extraction

Files:

- `k00_k14_full_hierarchy.tsv`
- `k00_k14_full_hierarchy.xlsx`

Columns:

- `章`
- `节`
- `类目代码`
- `类目名称`
- `亚目代码`
- `亚目名称`
- `诊断代码`
- `诊断名称`

This layer preserves the official Chinese hierarchy from the PDF.

### v1 cleaned master table

Files:

- `K00-K14_master_table_v1.tsv`
- `K00-K14_master_table_v1.xlsx`

Adds:

- `chapter_code`
- `parent_code`
- `is_subtype`
- `subtype_number`
- `code_level`

### English-column v1

Files:

- `K00-K14_master_table_v1_english_columns.tsv`
- `K00-K14_master_table_v1_english_columns.xlsx`

Renames columns to English snake_case. Chinese diagnosis text remains unchanged.

### v2 WHO English mapping layer

Files:

- `K00-K14_master_table_v2.tsv`
- `K00-K14_master_table_v2.xlsx`

Adds:

- `diagnosis_name_en`
- `category_name_en`
- `subcategory_name_en`
- `english_mapping_confidence`

This layer maps category and subcategory names to official WHO ICD-10 English terms where possible. Local China-specific rows inherit the closest WHO parent term.

### v3 semantic English master table

Files:

- `K00-K14_master_table_v3_semantic_en.tsv`
- `K00-K14_master_table_v3_semantic_en.xlsx`

Adds:

- `who_icd_code`
- `who_name_en`
- `structural_name_en`
- `semantic_name_en`
- `english_mapping_type`
- `semantic_source`

This is the current semantic English master table.

### v4 QA review layer

Files:

- `K00-K14_master_table_v4_QA.tsv`
- `K00-K14_master_table_v4_QA.xlsx`

Adds:

- `qa_status`
- `qa_reason`
- `qa_priority`

This layer flags rows that need manual review. It does not change v3 content.

## Validation Summary

Current row count:

```text
651
```

Validation checks passed through v4:

- no missing `diagnosis_code`
- no duplicate `diagnosis_code`
- all rows are within `K00-K14`
- subtype rows are preserved
- v3 existing columns remain unchanged in v4

## Rebuild Order

Run scripts in this order:

```bash
python3 scripts/extract_k00_k14.py
python3 scripts/build_master_table_v1.py
python3 scripts/rename_master_table_columns.py
python3 scripts/build_master_table_v2_who_english.py
python3 scripts/build_master_table_v3_semantic_en.py
python3 scripts/build_master_table_v4_QA.py
```

Notes:

- `extract_k00_k14.py` requires the source PDF to exist locally as `ICD-10医保2.0版.pdf`.
- Later scripts build from the previous TSV layer.
- The scripts write both TSV and XLSX outputs.

## Python Dependencies

The scripts use:

- `pdfplumber` for PDF table extraction
- `openpyxl` for XLSX output

Install if needed:

```bash
python3 -m pip install pdfplumber openpyxl
```

## WHO ICD-10 English Source

WHO-level English terms are based on the official WHO ICD-10 2019 browser:

<https://icd.who.int/browse10/2019/en#/K00-K14>

Important distinction:

- `diagnosis_code` is the China diagnosis code from the source PDF.
- `who_icd_code` is the closest official WHO ICD-10 code.
- China-specific local subtype codes such as `K00.100x001` are not treated as official WHO ICD-10 codes.

## Manual QA

Use:

- `K00-K14_master_table_v4_QA.xlsx`
- `K00-K14_v4_QA_review_instructions.md`

Start with:

```text
qa_status = NEEDS_REVIEW
english_mapping_type = LOCAL_SUBTYPE_TRANSLATION
```

These are China-specific local subtype rows and are the most important manual review group.

## Suggested Repository Structure Later

The current flat folder is acceptable for the first GitHub upload. Later, the project can be reorganized as:

```text
data/
  raw/
  intermediate/
  final/
docs/
scripts/
README.md
```

For now, keeping the files flat is simpler and less risky.
