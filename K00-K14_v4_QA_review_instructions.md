# K00-K14 v4 QA Review Instructions

## File To Review

Use:

- `K00-K14_master_table_v4_QA.xlsx`
- or `K00-K14_master_table_v4_QA.tsv`

The QA layer is for reviewing English semantic mappings only. It does not change the official Chinese ICD diagnosis codes, Chinese names, hierarchy, or WHO parent mappings.

## Review Goal

Confirm whether `semantic_name_en` is clinically acceptable for each Chinese diagnosis.

The main question is:

> Does `semantic_name_en` accurately represent `diagnosis_name_cn` using standard clinical or dental English?

## Recommended Review Order

1. Filter `qa_status = NEEDS_REVIEW`.
2. Review rows by `qa_priority`:
   - `HIGH` first, if any appear in later versions
   - `MEDIUM`
   - `LOW`
3. Within each priority, review by `english_mapping_type`:
   - `LOCAL_SUBTYPE_TRANSLATION`
   - `CLINICAL_TRANSLATION`
   - `WHO_PARENT_INHERITED`
   - `CLINICAL_STANDARD`
   - `WHO_EXACT`

## Columns To Compare

For each review row, compare:

- `diagnosis_code`
- `diagnosis_name_cn`
- `who_icd_code`
- `who_name_en`
- `structural_name_en`
- `semantic_name_en`
- `english_mapping_type`
- `english_mapping_confidence`
- `semantic_source`
- `qa_reason`

## What Each QA Reason Means

### LOW confidence English mapping

This usually means the row is a China-specific subtype or local extension code.

Check whether `semantic_name_en` is a medically accurate English term.

Examples:

| Chinese | Current semantic English |
|---|---|
| 少牙畸形 | Oligodontia |
| 先天缺牙 | Congenital absence of teeth |
| 舌牙痕 | Scalloped tongue |

If the English term is correct, set `qa_status = PASS`.

### Clinical translation requires review

This means the term was translated conservatively because no exact WHO ICD-10 English term exists.

Check against one of:

- recognized dental terminology
- SNOMED CT
- UMLS
- MeSH
- clinical dictionaries
- local hospital terminology standards

If the term is acceptable, keep or update `semantic_name_en` and set `qa_status = PASS`.

### Semantic source is clinical translation

This is similar to the previous reason. It means the English term did not come directly from WHO ICD-10 or a known dental terminology override.

Review carefully for wording, specificity, and clinical naturalness.

### Subtype semantic name matches structural parent term

This means the subtype English name is identical to the inherited parent term.

This can be acceptable when the subtype is essentially the same clinical concept as the parent.

Example:

| diagnosis_code | diagnosis_name_cn | semantic_name_en |
|---|---|---|
| K00.100x001 | 多生牙 | Supernumerary teeth |

If the subtype should be more specific, update `semantic_name_en`.

### Semantic English term appears generic for Chinese diagnosis

This means `semantic_name_en` may be too broad.

Examples of generic terms include:

- `Other diseases of tongue`
- `Other and unspecified lesions of oral mucosa`
- `Disease of jaws, unspecified`

If `diagnosis_name_cn` is more specific, replace `semantic_name_en` with a more precise clinical English term.

### Missing who_icd_code

This is a structural problem. It should be reviewed with high priority.

Expected action:

- fill the closest official WHO ICD-10 code
- do not force China-specific `x001/x002` codes into WHO format

### Missing parent_code

This is a structural problem.

Expected action:

- check whether the row has a valid hierarchy parent
- do not change `diagnosis_code`

### Subtype parent_code does not exist

This is a structural problem for subtype rows.

Expected action:

- verify that the parent diagnosis row exists in the table
- do not drop the subtype row

## What Not To Change

Do not change:

- `diagnosis_code`
- `diagnosis_name_cn`
- `chapter`
- `section`
- `category_code`
- `category_name_cn`
- `subcategory_code`
- `subcategory_name_cn`
- `chapter_code`
- `parent_code`, unless fixing a confirmed hierarchy error
- `who_icd_code`, unless fixing a confirmed WHO parent-code error
- `who_name_en`, unless fixing a confirmed WHO parent-code error
- `structural_name_en`

## Columns You May Change During QA

Usually only change:

- `semantic_name_en`
- `english_mapping_type`
- `english_mapping_confidence`
- `semantic_source`
- `qa_status`
- `qa_reason`
- `qa_priority`

## Decision Rules

### Mark PASS

Use `PASS` when:

- `semantic_name_en` is clinically accurate
- the English wording is natural
- the term is specific enough for `diagnosis_name_cn`
- the WHO parent mapping is reasonable

### Keep NEEDS_REVIEW

Keep `NEEDS_REVIEW` when:

- the English term is plausible but not confirmed
- the term may need specialist dental review
- the row uses local China-specific terminology
- the current English term is too broad

### Raise Priority

Use `HIGH` when:

- parent hierarchy is missing or wrong
- WHO parent code is missing or wrong
- English term could materially misclassify the diagnosis

Use `MEDIUM` when:

- local subtype translation needs clinical confirmation
- term is understandable but may not be the best standard term

Use `LOW` when:

- the row is acceptable but could benefit from future cleanup
- the issue is wording preference rather than clinical meaning

## Suggested First Review Batch

Start with:

```text
qa_status = NEEDS_REVIEW
english_mapping_type = LOCAL_SUBTYPE_TRANSLATION
```

These are the China-specific subtype rows and are the most important manual review set.

Then review:

```text
english_mapping_type = CLINICAL_TRANSLATION
semantic_source = CLINICAL_TRANSLATION
```

These rows are conservative translations and should be checked against preferred clinical terminology sources.

