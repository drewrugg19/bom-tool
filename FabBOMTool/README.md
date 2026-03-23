# Fabrication BOM Tool — Web App

## Quick start

1. Install Python 3.11+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`
3. Start the development server: `python app.py`
4. Open `http://localhost:5000`

## Production deployment

Use a production WSGI server instead of Flask debug mode.

```bash
gunicorn -w 4 -b 0.0.0.0:5000 wsgi:app
```

### Runtime configuration

The application reads these optional environment variables:

- `FBT_HOST` — bind host for `python app.py`.
- `FBT_PORT` — bind port for `python app.py`.
- `FLASK_DEBUG` — set to `1` only for local debugging.

### Health check

A lightweight health endpoint is available at:

```text
GET /health
```

It reports application status, version, and whether a cached legend file is present.

## Security notes

- Admin passwords are stored as hashes in `data/settings.json`.
- The app keeps a mirrored `data/settings.backup.json` copy so updates or a damaged primary settings file do not wipe your saved configuration.
- The default admin password is still `FBT2026!` on first launch, so change it immediately in the Admin panel.
- Uploaded PDFs are written to a temporary uploads folder with unique filenames and removed after processing.
- Generated exports, cached data, and local databases are intentionally excluded from version control.

## Folder structure

```text
FabBOMTool/
├── app.py
├── wsgi.py
├── requirements.txt
├── core/
│   ├── logic.py
│   └── history.py
├── templates/
│   └── index.html
├── static/
│   ├── style.css
│   └── app.js
├── data/
│   └── .gitkeep
├── uploads/
│   └── .gitkeep
└── exports/
    └── .gitkeep
```

## Legend file behavior

The app uses the first available legend source in this order:

1. `legend_embedded.py`
2. `data/Legend.cache.json`
3. `Legend.xlsx`

You can upload `Legend.cache.json` from the Admin tab.

## Automated tests

Run the regression suite with:

```bash
python -m unittest discover -s ../tests -v
```
