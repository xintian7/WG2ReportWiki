# WG2ReportWiki

## Precompute LLM term summaries

The Streamlit app displays **LLM Summary**.

Canonical command:

```bash
/opt/anaconda3/envs/tsu/bin/python script/reports/generate_term_usage_summaries.py
```

## Script entrypoints

Use domain-specific canonical scripts:

- Converters
	- `script/converters/convert_to_md.py`
- Glossary pipelines
	- `script/glossary/merge_ar6_ar7sod_glossaries.py`
	- `script/glossary/ar6_fgd_glossary_to_xlsx.py`
	- `script/glossary/ar7_srcities_sod_annex_i_glossary_to_xlsx.py`
	- `script/glossary/glossary_md_to_xlsx.py`
	- `script/glossary/build_glossary_network.py`
- Report/review generation
	- `script/reports/reconstruct_srcities_report.py`
	- `script/reports/export_srcities_html.py`
	- `script/reports/generate_term_usage_summaries.py`
	- `script/reports/extract_executive_summaries.py`
- App/runtime
	- `script/app/srcities_streamlit_app.py`
	- `script/app/encrypt_srsod.py`
	- `script/app/env_loader.py`
- Integrations
	- `script/integrations/write2notion.py`

## Legacy command note

Phase 3 cleanup removed old top-level wrapper scripts in `script/`.

If you have existing automation using old paths such as `script/reconstruct_srcities_report.py` or `script/srcities_streamlit_app.py`, update them to the canonical paths listed above.
