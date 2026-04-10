# AWR Review App (Sprint 1 + Step 1)

## Implemented
- JSON upload and validation (`instances` required)
- Canonical model (`cohort -> database -> instance`)
- Two selection contexts (`Web` and `PPT`) with link/unlink toggle
- Scope filters (cohorts, databases, instances)
- Metric catalog with grouped checkboxes
  - Core summary metrics
  - Dynamic `database_statistics` metrics
  - Optional advanced panels (`top_sql`, `segment_io`)
- Global KPI summary cards
- Cohort overview table with drilldown
- Cohort bar chart (metric selector)
- Database statistics summary table (p95/p99/max)
- Top SQL table (filtered by selected DBs)
- Segment IO table (filtered by selected DBs)
- Drilldown to database and instance detail tables

## Open
The app can be opened directly from:
`/Users/airtonalmeida/Documents/codex/AWR Review/index.html`

For template-native PPT export from the app button, run the integrated server:

```bash
/Users/airtonalmeida/Documents/codex/AWR\ Review/.venv/bin/python \
  /Users/airtonalmeida/Documents/codex/AWR\ Review/app_server.py \
  --host 127.0.0.1 --port 8080 \
  --root /Users/airtonalmeida/Documents/codex/AWR\ Review \
  --template /Users/airtonalmeida/Downloads/Oracle_PPT-template_FY26.potx
```

Then open:
`http://localhost:8080`

Note: the UI is configured to require template-native export. If `/api/export-template` is unavailable, export is blocked instead of falling back to the browser-based exporter.

## Input sample
`/Users/airtonalmeida/Documents/AWR/Alelo/awr_miner_occa/awr_occa_upload_data.json`

## Template-Native PPT Export (Oracle `.potx`)
To generate a deck that uses Oracle template layouts (white/light slides) and native editable PowerPoint charts:

```bash
/Users/airtonalmeida/Documents/codex/AWR\ Review/.venv/bin/python \
  /Users/airtonalmeida/Documents/codex/AWR\ Review/template_export.py \
  --input /Users/airtonalmeida/Documents/AWR/Alelo/awr_miner_occa/awr_occa_upload_data.json \
  --template /Users/airtonalmeida/Downloads/Oracle_PPT-template_FY26.potx \
  --output /tmp/AWR_template_native_test.pptx
```

You can also pass a previously saved report JSON (the app `Save` output) to preserve PPT scope selections.
