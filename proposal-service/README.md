# Blue Lime Proposal Service

A FastAPI microservice that turns Blue Lime Excel proposal templates into
polished, branded PDF proposals. Designed to be called from the Blue Lime
Proposal Studio (Lovable + Supabase) but works as a standalone API too.

## What this does

1. Accepts a Blue Lime proposal Excel file (`.xlsx`) via POST
2. Parses the template (Summary, Premium Summary, Authorization, etc. tabs)
3. Renders a multi-page PDF with the Blue Lime brand (cover, premium
   comparison, coverage detail, SOV, authorization, disclosures, team)
4. Returns the PDF bytes

## Project layout

```
proposal-service/
├── app/
│   ├── __init__.py
│   ├── main.py                  # FastAPI routes + auth + error handling
│   ├── excel_parser.py          # Reads Blue Lime Excel templates → dict
│   └── proposal_generator.py    # ReportLab layout engine
├── brand/
│   ├── logo.png                 # Real Blue Lime logo (transparent PNG)
│   ├── watermark.png            # Lime-pattern banner
│   ├── cover_bg.png             # Tiled lime-pattern full page
│   └── headshots/               # Team headshots (briana.png, carol.png, ...)
├── requirements.txt
├── render.yaml                  # One-click Render deploy
├── .gitignore
└── README.md
```

## Local development

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run with auth disabled (for local testing only)
uvicorn app.main:app --reload --port 8000

# Or run with auth enabled
PROPOSAL_API_KEY=dev-secret uvicorn app.main:app --reload --port 8000
```

Then test:

```bash
curl -X POST http://localhost:8000/generate-proposal \
  -H "X-API-Key: dev-secret" \
  -F "file=@26-27_Proposal_Haven_at_Keith_Harrow.xlsx" \
  -o output.pdf
```

## Deployment to Render

1. Push this directory to a GitHub repo (e.g., `bluelime/proposal-service`).
2. In Render, click **New → Blueprint** and connect the repo.
3. Render reads `render.yaml`, creates the web service, and generates a
   strong `PROPOSAL_API_KEY` automatically.
4. After deploy, copy the URL (something like
   `https://bluelime-proposal-service.onrender.com`) and the
   `PROPOSAL_API_KEY` value (visible under "Environment" in Render).
5. Paste both into Lovable as Supabase secrets:
   - `PROPOSAL_SERVICE_URL`
   - `PROPOSAL_API_KEY`

## API reference

### `GET /health`

Returns `{"status": "ok", "time": "...", "auth": "enabled"}`. Used by
Render's health check.

### `POST /generate-proposal`

**Headers:**
- `X-API-Key: <secret>` — required when `PROPOSAL_API_KEY` env is set.

**Body:** `multipart/form-data`
- `file` (required): `.xlsx` or `.xlsm` file, max 20 MB.

**Responses:**
- `200`: `application/pdf` with `Content-Disposition: attachment;
  filename="<short-name>_Proposal.pdf"`.
- `400`: Wrong file type / empty file.
- `401`: Bad or missing API key.
- `413`: File too large.
- `422`: Excel could not be parsed (invalid template structure).
- `500`: Generator error.

The PDF response also includes a non-standard `X-Generated-For` header
showing the parsed client short name — useful for logging on the caller.

## How the parser handles template variation

The parser is tolerant of:

- Missing cells (returns sensible defaults like "Not Included")
- Placeholder values (`XXXX`, `XXXXX`, `N/A`, blank) → treated as absent
- Numeric vs. string cells for the same field
- Either `Cover Page` or `Summary` tab as the source of the client name

Cell coordinates are documented in `app/excel_parser.py`. If your team's
template structure changes, that's the file to update.

## Adding a new template variant

If you build proposals from a different Excel layout (e.g., for condo
master policies or coastal accounts):

1. Send a sample file to the developer maintaining this service.
2. They'll either extend `parse_excel()` to detect the variant or add a
   new parser function and route based on a marker cell.

## Updating brand assets

To replace the logo, headshots, or watermark, drop the new PNG files into
`brand/` (or `brand/headshots/`) with the same filenames and redeploy.

Headshot file slugs must match the lowercase first name of each team
member (e.g., `briana.png`, `david.png`). To add a new team member:

1. Add their headshot to `brand/headshots/<slug>.png`.
2. Add the row to `_team_page()` in `app/proposal_generator.py`.
3. Add their `<slug>` mapping in the `SLUG_MAP` dict in
   `app/excel_parser.py`.
