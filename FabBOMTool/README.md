# Fabrication BOM Tool — Web App

## Quick Start

1. Make sure Python is installed on your PC
2. Double-click **START.bat**
3. Open your browser to **http://localhost:5000**

That's it. The START.bat installs dependencies automatically on first run.

---

## Folder Structure

```
FabBOMTool/
├── app.py              ← Flask server (run this)
├── START.bat           ← Double-click to launch on Windows
├── requirements.txt    ← Python dependencies
│
├── core/
│   ├── logic.py        ← All PDF parsing, classification, multiplier logic
│   └── history.py      ← Run history (SQLite)
│
├── templates/
│   └── index.html      ← Main web page
│
├── static/
│   ├── style.css       ← Styles
│   └── app.js          ← Browser interactions
│
├── data/
│   ├── settings.json   ← App settings (auto-created)
│   └── history.db      ← Run history database (auto-created)
│
├── uploads/            ← Temp folder for uploaded PDFs (auto-cleared)
└── exports/            ← Excel exports saved here
```

---

## Legend file

The app uses your legend for fitting classification in this order:

1. `legend_embedded.py` placed next to `app.py` ← preferred
2. `data/Legend.cache.json` ← upload via Admin panel
3. `Legend.xlsx` placed next to `app.py` ← legacy fallback

---

## Letting teammates connect

By default the server listens on all network interfaces (`0.0.0.0`).
Teammates on your network can open `http://YOUR-PC-IP:5000` in their browser.

To find your PC's IP: open Command Prompt and type `ipconfig`
Look for the IPv4 Address under your network adapter.

If they can't connect, it may be a Windows Firewall rule.
To allow it: Windows Firewall → Allow an app → add Python.

---

## Updates

To update the app, replace any file and restart the server (Ctrl+C in the terminal, then run START.bat again).
Changes take effect immediately for all users on next page refresh.

---

## Default admin password

`FBT2026!`

Change it in the Admin panel after first login.
