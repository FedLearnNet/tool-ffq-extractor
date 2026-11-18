# FFQ Extractor - Tool

FL-Net **EXTRACTOR** tool that converts **food frequency questionnaire (FFQ)** Excel workbooks into a flat CSV mapped to the microbiome **Nutrition** schema.

Supports:

- a single filled `.xlsx`
- a `.zip` of many per-patient FFQ Excel files (designed for **100k+** patients)

Follows the platform [Create a New App](https://dev.federated-learning.net/documentation/docs/tool-dev/create-tool) flow (`pyfedappwrap`, `app.yml`, Docker).

## Platform setup

1. In the FL-Net UI, open `/app` → **New App** / **New Tool**:
   - Name: `FFQ Extractor - Tool`
   - Slug: `tool-ffq-extractor`
   - Type: `EXTRACTOR`
2. Copy the assigned **App ID** into `.env` (see `.env.example`).
3. Push this repo and run the tool **build pipeline** to publish the Docker image.

## Local install

```bash
cd ffq-extractor
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install -e ".[dev]"
```

## FL-Net app entrypoint

```bash
python main.py
```

Contract (`app.yml` / `config.py`):

| | Name | Type |
|--|------|------|
| Hyperparam | `workers` | process-pool size for ZIP (default `4`) |
| Hyperparam | `sheet` | worksheet name; empty = first sheet |
| Hyperparam | `date_cell` | questionnaire date cell (default `A1`) |
| Hyperparam | `patient_id_cell` | patient ID cell (default `A3`) |
| Input | `ffq_archive` | `PATH` (`.xlsx` or `.zip`) |
| Output | `nutrition_csv` | `CSV` |

## CLI (offline / nurses)

```bash
ffq-extract filled.xlsx -o nutrition.csv
ffq-extract patients.zip -o nutrition.csv --workers 8
```

ZIP batch behavior:

- walks nested folders;
- continues on per-file failures; logs errors to console / Docker logs
- streams part CSVs and merges as tasks complete
- rejects oversized ZIP members (default 50 MiB uncompressed)
- exit code `1` if any member failed or none succeeded

## Output columns

| Column | Notes |
|--------|--------|
| `Unique Patient ID` | value from `patient_id_cell` (default `A3`); error if blank |
| `Questionnaire date` | value from `date_cell` (default `A1`); ISO date when Excel stores a date; error if blank |
| `Portion size` | text after ` - ` in section headers when present |
| `Frequency consumption` | schema codes (`1/D`, `<1/M`, `6+/D`, …) |
| `Food group` / `Food subgroup` | translated English values from the mapping file |
| `Food item` | FFQ item label |

## Frequency codes

| Excel header | CSV code |
|--------------|----------|
| Less than once a month | `<1/M` |
| Once a month | `1/M` |
| Once a week | `1/W` |
| Twice a week | `2/W` |
| 5 times a week | `5/W` |
| Once a day | `1/D` |
| Twice a day | `2/D` |
| 4 times a day | `4/D` |
| 6 or more times a day | `6+/D` |
| Never | row omitted |

Accepted marks: `X`, `x`, `1`, `✓`, `✔`, `Y`/`yes`.

## Connector

Upload `nutrition_csv` via the FILE connector:

`meta/deployment/example-data/connectors/2026-08-10microbiome_nutrition_ffq_connector.json`

Requires a harmonized microbiome schema with the **Nutrition** group (`2026-08-10`+). Schema node **`Food item `** includes a trailing space; the connector mapping preserves it.

## Development

```bash
python -m pytest -q
```

## License

MIT
