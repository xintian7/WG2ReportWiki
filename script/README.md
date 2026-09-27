# Script Organization

The script folder is domain-organized. Use canonical entrypoints in subfolders.

## Canonical locations

- Converters
  - script/converters/convert_to_md.py

- Glossary pipelines
  - script/glossary/merge_ar6_ar7sod_glossaries.py
  - script/glossary/ar6_fgd_glossary_to_xlsx.py
  - script/glossary/ar7_srcities_sod_annex_i_glossary_to_xlsx.py
  - script/glossary/glossary_md_to_xlsx.py
  - script/glossary/build_glossary_network.py

- Report/review generation
  - script/reports/reconstruct_srcities_report.py
  - script/reports/export_srcities_html.py
  - script/reports/generate_term_usage_summaries.py
  - script/reports/extract_executive_summaries.py

- App/runtime
  - script/app/srcities_streamlit_app.py
  - script/app/encrypt_srsod.py
  - script/app/env_loader.py

- Integrations
  - script/integrations/write2notion.py

## Legacy wrappers

Phase 3 cleanup removed top-level compatibility wrappers.

- Use canonical paths under script/converters, script/glossary, script/reports, script/app, and script/integrations.

## Recommended usage

Use script/converters/convert_to_md.py for all new conversion workflows.

- DOCX -> Markdown
  - /opt/anaconda3/envs/tsu/bin/python script/converters/convert_to_md.py --input <file.docx> --output <file.md>

- PDF -> Markdown
  - /opt/anaconda3/envs/tsu/bin/python script/converters/convert_to_md.py --input <file.pdf> --output <file.md>

## Additional cleanup already applied

- data/iamges4UG renamed to data/images4UG
- temporary exports grouped under data/export/tmp/
