# AWR Review App

Web application + PowerPoint generator for reviewing Oracle AWR/OCCA JSON outputs and producing executive-ready reports.

## What This Application Does

- Loads AWR/OCCA JSON files and builds an interactive analysis workspace.
- Organizes data by:
  - `Cohort`
  - `Database`
  - `Instance`
- Shows global and drill-down analytics (cohort and instance levels).
- Lets users select/deselect:
  - Cohorts, databases, instances
  - Metrics to display
  - PPT slides to export
- Exports a PowerPoint report using an Oracle template (`.potx`) with:
  - Title/divider/content/closing structure
  - Summary, cohort, and instance sections
  - Native PPT charts + Python-rendered boxplots where needed

## Main Features

- Global dashboard with KPIs (DBs, hosts, instances, vCPU, memory, storage).
- CPU and memory visualizations with drill-down navigation.
- Cohort and instance detail pages.
- Sortable tables for major datasets.
- Save/Load report state (selection + layout choices).
- Template-native PPT export via local API server.

## Project Files

- `/Users/airtonalmeida/Documents/codex/AWR Review/index.html` - UI shell
- `/Users/airtonalmeida/Documents/codex/AWR Review/app.js` - frontend logic, charts, interactions
- `/Users/airtonalmeida/Documents/codex/AWR Review/styles.css` - styling
- `/Users/airtonalmeida/Documents/codex/AWR Review/app_server.py` - local server + export API
- `/Users/airtonalmeida/Documents/codex/AWR Review/template_export.py` - template-based PPT generator

## Architecture

```mermaid
flowchart LR
  U["User (Browser)"] --> FE["Frontend UI (index.html + app.js + styles.css)"]
  FE -->|Load JSON| J["AWR/OCCA JSON Data Model"]
  FE -->|Export PPT (HTTP)| API["Local API Server (app_server.py)"]
  API -->|Build slides| EXP["Template Export Engine (template_export.py)"]
  EXP -->|Read template| T["Oracle .potx Template"]
  EXP -->|Generate| P[".pptx Report"]
```

### Runtime Flow

1. Frontend loads JSON and builds normalized structures for cohorts, databases, instances, and metrics.
2. User configures filters and slide selection in the UI.
3. On `Export PPT`, frontend sends report state to `app_server.py`.
4. Server calls `template_export.py` to merge selected data into Oracle template layouts.
5. Generated `.pptx` is streamed back to browser for download.

## How To Run

From project root:

```bash
cd "/Users/airtonalmeida/Documents/codex/AWR Review"
```

Start the app server (required for template-native PPT export):

```bash
./.venv/bin/python app_server.py \
  --host 127.0.0.1 --port 8080 \
  --root "/Users/airtonalmeida/Documents/codex/AWR Review" \
  --template "/Users/airtonalmeida/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/AWR Template.potx"
```

Open in browser:

- `http://127.0.0.1:8080`

## Typical Workflow

1. Load JSON (`Load JSON`)
2. Optionally load/update template (`Load Template`)
3. Review global/cohort/instance analytics
4. Adjust scope + metric selections
5. Choose PPT slides in `PPT Storyboard`
6. Click `Export PPT`

## Notes

- Export is configured to prefer template-native generation.
- If export server/template is unavailable, the app will show a clear error.
- Recommended input example:
  - `/Users/airtonalmeida/Documents/AWR/Alelo/awr_miner_occa/awr_occa_upload_data.json`
