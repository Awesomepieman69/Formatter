# Formatter (MLA 9th Edition DOCX Formatter)

Web app that reformats essays into MLA 9th edition using Flask + python-docx.

## Features
- MLA heading handling (replace or preserve existing heading)
- Running header (LastName + page number)
- Title/body/Works Cited formatting
- Empty-line cleanup for MLA spacing
- AI-assisted structure detection and compliance warnings (optional)
- Works Cited rewrite support (optional)

## Local Run
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. (Optional) set OpenAI key in `.env`:
   ```bash
   OPENAI_API_KEY=your_key_here
   ```
3. Start app:
   ```bash
   python3 app.py
   ```
4. Open [http://localhost:8080](http://localhost:8080)

## Deploy Publicly (Render)
This repo includes `render.yaml` for one-click deployment.

1. Push this repo to GitHub.
2. In Render: New + -> Blueprint -> select this repo.
3. Add environment variable `OPENAI_API_KEY` (optional, for AI features).
4. Deploy.

Your app will be public at `https://<your-service>.onrender.com`.

## File Overview
- `app.py` - Flask server/routes
- `mla_formatter.py` - MLA formatting pipeline
- `templates/index.html` - frontend UI
