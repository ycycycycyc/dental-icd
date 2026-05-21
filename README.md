# Dental ICD-10 K00-K14 Chinese-English Master Table

This repository contains a structured Chinese-English master table for ICD-10 `K00-K14`, covering diseases of the oral cavity, salivary glands, jaws, teeth, gingiva, tongue, and related oral structures.

The project starts from the Chinese `ICD-10医保2.0版` hierarchy, preserves China diagnosis codes and local subtype codes, maps each row to the closest WHO ICD-10 English parent code, and adds a semantic English diagnosis name for lookup and review.

## Main Outputs

| File | Description |
|---|---|
| `K00-K14_master_table_v3_semantic_en.tsv` | Current semantic English master table |
| `K00-K14_master_table_v3_semantic_en.xlsx` | Excel version of the semantic master table |
| `K00-K14_master_table_v4_QA.tsv` | QA review layer with review flags |
| `K00-K14_master_table_v4_QA.xlsx` | Excel version of the QA review layer |
| `K00-K14_v3_semantic_master_summary_and_data_dictionary.md` | Data dictionary and project summary |
| `K00-K14_v4_QA_review_instructions.md` | Manual review guide for QA rows |

## What This Data Is For

This table is designed for:

- bilingual dental ICD lookup
- structured terminology review
- mapping China ICD-10 local diagnosis codes to closest WHO ICD-10 parent codes
- building search tools for Chinese and English oral diagnosis terms
- preparing a manually reviewed semantic terminology layer

It is not a treatment, billing, reimbursement, or clinical decision-support table.

## Side Project: Dental ICD Lookup

This repository also includes a small static lookup app:

- [K00-K14 Dental ICD Lookup / 口腔 ICD-10 双语诊断编码查询](projects/dental-icd-lookup/)

The lookup app is a single-page HTML/CSS/JavaScript project that can run on GitHub Pages. It loads the v3 semantic table and supports search by:

- partial China diagnosis code, such as `K02` or `K02.0`
- Chinese diagnosis terms, such as `龋` or `牙髓炎`
- English terms, such as `caries` or `pulpitis`
- WHO ICD code and WHO English parent term

Each result shows the China diagnosis code, Chinese diagnosis name, semantic English diagnosis name, WHO ICD code, WHO English name, mapping confidence, and mapping type. Rows can be expanded to inspect hierarchy and source details.

Expected GitHub Pages path:

```text
[[https://<your-username>.github.io/<repo-name>/projects/dental-icd-lookup/](https://ycycycycyc.github.io/dental-icd/projects/dental-icd-lookup/)]
```

## Repository Layout

```text
.
├── README.md
├── K00-K14_master_table_v3_semantic_en.tsv
├── K00-K14_master_table_v3_semantic_en.xlsx
├── K00-K14_master_table_v4_QA.tsv
├── K00-K14_master_table_v4_QA.xlsx
├── K00-K14_v3_semantic_master_summary_and_data_dictionary.md
├── K00-K14_v4_QA_review_instructions.md
├── scripts/
│   ├── README.md
│   ├── extract_k00_k14.py
│   ├── build_master_table_v1.py
│   ├── rename_master_table_columns.py
│   ├── build_master_table_v2_who_english.py
│   ├── build_master_table_v3_semantic_en.py
│   └── build_master_table_v4_QA.py
└── projects/
    └── dental-icd-lookup/
        ├── index.html
        ├── styles.css
        ├── app.js
        └── data/
            └── K00-K14_master_table_v3_semantic_en.tsv
```

## Data Layers

### Raw hierarchy extraction

The raw extraction preserves the original Chinese hierarchy:

- chapter
- section
- category
- subcategory
- diagnosis code
- diagnosis name

It keeps subtype codes such as `K00.100x001` and does not collapse them into parent rows.

### v1 cleaned hierarchy

The v1 layer adds structural fields:

- `chapter_code`
- `parent_code`
- `is_subtype`
- `subtype_number`
- `code_level`

### v2 WHO English mapping

The v2 layer adds official WHO ICD-10 English labels where available:

- `diagnosis_name_en`
- `category_name_en`
- `subcategory_name_en`
- `english_mapping_confidence`

This layer is intentionally structural. Local China-specific diagnosis rows inherit the closest WHO ICD-10 parent term.

### v3 semantic English master table

The v3 layer is the current master table. It adds:

- `who_icd_code`
- `who_name_en`
- `structural_name_en`
- `semantic_name_en`
- `english_mapping_type`
- `semantic_source`

`semantic_name_en` is the best current English clinical label for the row. It may be an official WHO ICD-10 term, a standard dental term, or a conservative clinical translation.

Examples:

| diagnosis_code | diagnosis_name_cn | who_icd_code | structural_name_en | semantic_name_en |
|---|---|---|---|---|
| `K00.000x003` | 少牙畸形 | `K00.0` | Anodontia | Oligodontia |
| `K00.000x004` | 先天缺牙 | `K00.0` | Anodontia | Congenital absence of teeth |
| `K04.007` | 不可逆性牙髓炎 | `K04.0` | Pulpitis | Irreversible pulpitis |
| `K14.800x010` | 舌牙痕 | `K14.8` | Other diseases of tongue | Scalloped tongue |

### v4 QA layer

The v4 layer appends manual review fields:

- `qa_status`
- `qa_reason`
- `qa_priority`

Rows marked `NEEDS_REVIEW` should be reviewed before treating the semantic English term as final.

## Key Concepts

### China diagnosis code

`diagnosis_code` is the China ICD diagnosis code from the source hierarchy.

Examples:

- `K04.000`
- `K04.007`
- `K00.100x001`

China-specific local subtype codes are preserved as-is.

### WHO ICD code

`who_icd_code` is the closest official WHO ICD-10 code.

Example:

```text
diagnosis_code = K00.100x001
diagnosis_name_cn = 多生牙
who_icd_code = K00.1
who_name_en = Supernumerary teeth
```

The local subtype code is not forced into WHO format.

### Structural vs semantic English

`structural_name_en` is inherited from the WHO parent mapping.

`semantic_name_en` is the clinically refined English diagnosis label.

Example:

```text
diagnosis_name_cn = 少牙畸形
structural_name_en = Anodontia
semantic_name_en = Oligodontia
```

## Validation

Current validation status:

```text
Total rows: 651
Missing diagnosis_code: 0
Duplicate diagnosis_code: 0
Rows outside K00-K14: 0
Subtype rows preserved: yes
```

v4 QA status:

```text
PASS: 356
NEEDS_REVIEW: 295
```

## Reproducing The Tables

The scripts are designed to run in order:

```bash
python3 scripts/extract_k00_k14.py
python3 scripts/build_master_table_v1.py
python3 scripts/rename_master_table_columns.py
python3 scripts/build_master_table_v2_who_english.py
python3 scripts/build_master_table_v3_semantic_en.py
python3 scripts/build_master_table_v4_QA.py
```

The raw PDF source is not included in this repository. To reproduce the extraction step, place the source PDF in the project root as:

```text
ICD-10医保2.0版.pdf
```

The later scripts can be inspected independently because the mapping logic is encoded in the scripts.

## Script Summary

| Script | Purpose |
|---|---|
| `scripts/extract_k00_k14.py` | Extracts K00-K14 rows from the source PDF while preserving hierarchy |
| `scripts/build_master_table_v1.py` | Adds structural hierarchy fields such as parent code and subtype flags |
| `scripts/rename_master_table_columns.py` | Renames Chinese headers to English snake_case |
| `scripts/build_master_table_v2_who_english.py` | Adds WHO ICD-10 English category/subcategory mappings |
| `scripts/build_master_table_v3_semantic_en.py` | Adds semantic English diagnosis names |
| `scripts/build_master_table_v4_QA.py` | Adds QA review flags |

## Dependencies

Python packages:

- `pdfplumber`
- `openpyxl`

Install with:

```bash
python3 -m pip install pdfplumber openpyxl
```

## Sources

WHO ICD-10 English terms are based on the official WHO ICD-10 2019 browser:

<https://icd.who.int/browse10/2019/en#/K00-K14>

The China diagnosis codes and Chinese diagnosis names are derived from the `ICD-10医保2.0版` PDF source.

Original Chinese ICD-10 Source
This repository does not redistribute the original `ICD-10医保2.0版.pdf` source file.
The China diagnosis codes and Chinese diagnosis names were derived from the public `ICD-10医保2.0版` source available through the official 国家医保信息业务编码标准数据库动态维护 platform.
To reproduce the extraction step, obtain the PDF from the official source, place it in the project root as:

```text
ICD-10医保2.0版.pdf

## Limitations

- The source PDF is not included.
- China-specific local subtype codes are not official WHO ICD-10 codes.
- Some semantic English names are conservative clinical translations and require manual review.
- Rows marked `NEEDS_REVIEW` in v4 should not be treated as fully QA-approved.

## Recommended Next Work

- Review all `NEEDS_REVIEW` rows in `K00-K14_master_table_v4_QA.xlsx`.
- Move confirmed rows from `NEEDS_REVIEW` to `PASS` after terminology review.
- Add a changelog for manual terminology corrections.
- Publish the lookup app with GitHub Pages and update the README link with the live URL.
