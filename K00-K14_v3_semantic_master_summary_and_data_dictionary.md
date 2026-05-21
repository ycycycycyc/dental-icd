# K00-K14 Semantic English Master Table Summary

## Current Master Table

The current semantic English master table is:

- `K00-K14_master_table_v3_semantic_en.tsv`
- `K00-K14_master_table_v3_semantic_en.xlsx`

This table should be treated as the frozen semantic English master layer before QA review.

## What Has Been Done So Far

### 1. Raw official extraction from PDF

Source:

- `ICD-10医保2.0版.pdf`

Output:

- `k00_k14_full_hierarchy.tsv`
- `k00_k14_full_hierarchy.xlsx`

What was done:

- Extracted only diagnosis rows where the diagnosis code starts with `K00` through `K14`.
- Preserved the original Chinese ICD hierarchy from the PDF.
- Preserved all subtype rows, including local extension codes such as `K00.100x001`.
- Preserved duplicate-looking Chinese clinical terms unless the diagnosis code itself was duplicated.
- Preserved the original Chinese text from the PDF.

Result:

- 651 rows
- no missing diagnosis codes
- no rows outside K00-K14
- no duplicate diagnosis codes

### 2. Cleaned structured master table v1

Output:

- `K00-K14_master_table_v1.tsv`
- `K00-K14_master_table_v1.xlsx`

What was added:

- `chapter_code`
- `parent_code`
- `is_subtype`
- `subtype_number`
- `code_level`

Purpose:

- Create a structured Chinese ICD hierarchy master table with stable parent-child fields.

### 3. English snake_case column version

Output:

- `K00-K14_master_table_v1_english_columns.tsv`
- `K00-K14_master_table_v1_english_columns.xlsx`

What was done:

- Renamed the original Chinese column headers to English snake_case.
- Did not translate Chinese diagnosis names.
- Did not change row content.

### 4. WHO English mapping layer v2

Output:

- `K00-K14_master_table_v2.tsv`
- `K00-K14_master_table_v2.xlsx`

What was added:

- `diagnosis_name_en`
- `category_name_en`
- `subcategory_name_en`
- `english_mapping_confidence`

Mapping rule:

- WHO ICD-10 English terms were used for official category and subcategory codes.
- Canonical ICD rows such as `K04.000` inherited the official WHO term for `K04.0`.
- China-specific extension rows inherited the closest WHO parent term.
- Subtype rows were marked as lower confidence because they are not official WHO ICD-10 codes.

Important limitation:

- v2 was a structural WHO parent mapping layer, not a full semantic clinical translation layer.

### 5. Semantic English refinement layer v3

Output:

- `K00-K14_master_table_v3_semantic_en.tsv`
- `K00-K14_master_table_v3_semantic_en.xlsx`

What was added:

- `who_icd_code`
- `who_name_en`
- `structural_name_en`
- `semantic_name_en`
- `english_mapping_type`
- `semantic_source`

Purpose:

- Keep the official WHO parent mapping.
- Add a more clinically meaningful English term for each diagnosis row.
- Preserve China-specific diagnosis codes and Chinese diagnosis names.
- Avoid pretending China-specific subtype codes are official WHO codes.

Examples:

| diagnosis_name_cn | structural_name_en | semantic_name_en |
|---|---|---|
| 少牙畸形 | Anodontia | Oligodontia |
| 先天缺牙 | Anodontia | Congenital absence of teeth |
| 不可逆性牙髓炎 | Pulpitis | Irreversible pulpitis |
| 急性龈乳头炎 | Acute gingivitis | Acute gingival papillitis |
| 舌牙痕 | Other diseases of tongue | Scalloped tongue |

### 6. QA layer v4

Output:

- `K00-K14_master_table_v4_QA.tsv`
- `K00-K14_master_table_v4_QA.xlsx`

What was added:

- `qa_status`
- `qa_reason`
- `qa_priority`

Purpose:

- Identify rows needing manual review.
- Do not change v3 content.
- Provide a practical review layer for semantic English QA.

## v3 Row Count And Validation

For `K00-K14_master_table_v3_semantic_en`:

- total rows: 651
- missing diagnosis codes: 0
- duplicate diagnosis codes: 0
- rows outside K00-K14: 0
- subtype rows preserved: yes
- original v2 columns preserved: yes

## v3 Column Dictionary

### `chapter`

Chinese ICD chapter hierarchy label from the original PDF.

Example:

`11 K00-K93 消化系统疾病`

### `section`

Chinese ICD section/block label from the original PDF.

Example:

`K00-K14 口腔、涎腺和颌疾病`

### `category_code`

Three-character ICD category code.

Example:

`K04`

### `category_name_cn`

Original Chinese category name from the PDF.

Example:

`牙髓和根尖周组织疾病`

### `subcategory_code`

WHO ICD-10 subcategory code.

Example:

`K04.0`

### `subcategory_name_cn`

Original Chinese subcategory name from the PDF.

Example:

`牙髓炎`

### `diagnosis_code`

China ICD diagnosis code from the PDF.

This may be:

- a canonical local diagnosis code such as `K04.000`
- a local extension code such as `K04.001`
- a local subtype code such as `K00.100x001`

Do not modify this field.

### `diagnosis_name_cn`

Original Chinese diagnosis name from the PDF.

Do not modify this field unless correcting a confirmed PDF extraction error.

### `chapter_code`

First three characters of `diagnosis_code`.

Example:

`K04`

### `parent_code`

Parent diagnosis code used for hierarchy.

Rules:

- for subtype codes such as `K00.100x001`, parent is `K00.100`
- for non-subtype codes, parent is the subcategory code, such as `K00.1`

### `is_subtype`

Boolean subtype flag.

Values:

- `TRUE` if `diagnosis_code` contains `x`
- `FALSE` otherwise

### `subtype_number`

Subtype suffix number extracted from local `x` codes.

Example:

`K00.100x001` has subtype number `001`.

Blank for non-subtype rows.

### `code_level`

Structural level of the row.

Allowed values used so far:

- `diagnosis`
- `subtype`

Earlier logic also reserved:

- `category`
- `subcategory`

But the current K00-K14 extracted rows are diagnosis/subtype-level rows.

### `diagnosis_name_en`

v2 structural English diagnosis name.

This is usually inherited from the closest WHO ICD-10 subcategory.

Example:

`K00.000x003 少牙畸形` has `diagnosis_name_en = Anodontia`.

This column is intentionally preserved from v2.

### `category_name_en`

Official WHO ICD-10 English category name for `category_code`.

Example:

`K04` -> `Diseases of pulp and periapical tissues`

### `subcategory_name_en`

Official WHO ICD-10 English subcategory name for `subcategory_code`.

Example:

`K04.0` -> `Pulpitis`

### `english_mapping_confidence`

Confidence flag from the v2 English mapping layer.

Allowed values:

- `HIGH`
- `MEDIUM`
- `LOW`

Meaning:

- `HIGH`: official WHO/ICD exact terminology exists at the mapped level
- `MEDIUM`: close standardized medical equivalent or inherited WHO parent term
- `LOW`: China-specific subtype/localized term requiring manual review

### `who_icd_code`

Closest official WHO ICD-10 code for the row.

Usually this is the WHO subcategory code.

Example:

`K00.100x001 多生牙` has `who_icd_code = K00.1`.

Important:

- Do not force China-specific `x001/x002` codes into WHO format.
- China-specific codes remain in `diagnosis_code`.

### `who_name_en`

Official WHO ICD-10 English name for `who_icd_code`.

Example:

`K00.1` -> `Supernumerary teeth`

### `structural_name_en`

Inherited English term from the v2 parent WHO/ICD mapping.

This is the structural parent mapping, not necessarily the best clinical semantic English name.

Example:

`少牙畸形` has `structural_name_en = Anodontia`.

### `semantic_name_en`

Refined clinical English diagnosis name.

Priority used:

1. Official WHO ICD-10 English term if an exact WHO-level match exists.
2. Standard dental or clinical terminology when available.
3. Conservative clinical English translation if no standard term was available.

Examples:

- `少牙畸形` -> `Oligodontia`
- `先天缺牙` -> `Congenital absence of teeth`
- `不可逆性牙髓炎` -> `Irreversible pulpitis`

This is the primary English diagnosis name for downstream semantic use, but it still needs QA for review-flagged rows.

### `english_mapping_type`

Type of English mapping used for `semantic_name_en`.

Allowed values:

- `WHO_EXACT`
- `WHO_PARENT_INHERITED`
- `CLINICAL_STANDARD`
- `LOCAL_SUBTYPE_TRANSLATION`
- `CLINICAL_TRANSLATION`

Meaning:

- `WHO_EXACT`: row matches an official WHO ICD-10 term at the mapped WHO code level
- `WHO_PARENT_INHERITED`: semantic English still uses the WHO parent term
- `CLINICAL_STANDARD`: refined to a recognized clinical/dental term
- `LOCAL_SUBTYPE_TRANSLATION`: China-specific subtype translated/refined clinically
- `CLINICAL_TRANSLATION`: conservative translation requiring review

### `semantic_source`

Source category for `semantic_name_en`.

Allowed values used so far:

- `WHO_ICD10`
- `DENTAL_TERMINOLOGY`
- `CLINICAL_TRANSLATION`

Allowed but not yet actively populated:

- `SNOMED_CT`
- `UMLS`
- `MESH`

Meaning:

- `WHO_ICD10`: official WHO ICD-10 English terminology
- `DENTAL_TERMINOLOGY`: standard dental/clinical terminology
- `CLINICAL_TRANSLATION`: conservative translation pending review

## How To Use v3

Use v3 when you need:

- the official Chinese ICD diagnosis hierarchy
- China diagnosis codes preserved
- WHO parent ICD mapping
- a semantic English diagnosis label
- a distinction between WHO exact terms and local subtype translations

Use v4 when you need:

- review status
- QA reasons
- priority for manual review

## Important Notes

- `diagnosis_code` is the authoritative China diagnosis code.
- `who_icd_code` is the closest official WHO ICD-10 code.
- `semantic_name_en` is the best current English clinical label.
- China-specific subtype rows are not official WHO codes.
- QA review should focus mainly on rows marked `NEEDS_REVIEW` in v4.

