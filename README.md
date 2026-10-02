# WSA Educational Test Series & Paper Builder

AI-powered exam paper generator, answer checker and test diagnostic tool
(English / Urdu / Bilingual). Built with Streamlit.

## Features
- Class-based setup: choose School (Class 1-10), College (Class 11-12), University or a
  competitive test (NTS, PPSC / FPSC, BPSC, CSS, CS / IT). Subject list, difficulty and
  time are filled in automatically.
- Professional papers: MCQs (A-D), short and long questions, marks, answer key
- Export: PDF (paper and answer key), Word (.docx), text
- Check student answers (text, PDF, DOCX or a photo of the answer sheet)
- Test diagnostic with an improvement plan
- AI providers: **Groq** (primary) with automatic fallback to **Gemini**
- No settings or keys are shown to users. Keys stay on the server.
- Built-in usage limits protect the shared free API quota

## Run locally
```bash
pip install -r requirements.txt
cp .env.example .env        # then put your real keys inside .env
streamlit run app.py
```

## API keys (important)
Keys are read only from Streamlit secrets or environment variables, never from the code
and never from the user interface.

| Name | Required | Where to get it |
|------|----------|-----------------|
| `GROQ_API_KEY` | yes | https://console.groq.com/keys |
| `GEMINI_API_KEY` | recommended (fallback and photo reading) | https://aistudio.google.com/apikey |

Optional settings: `DAILY_LIMIT`, `GROQ_MODEL`, `GROQ_VISION_MODEL`, `GEMINI_MODEL`, `DEBUG`.
Use the model settings when a provider retires a model (no code change needed).

**Never commit `.env` or `secrets.toml`.** If a key is ever pushed to GitHub, revoke it in the
provider console immediately and create a new one.

## Deploy (Streamlit Community Cloud)
1. Push this repo to GitHub (keep it private at first).
2. share.streamlit.io -> New app -> select the repo and `app.py`.
3. Advanced settings -> Secrets -> paste the content of `secrets.toml.example` with your real keys.

## Roadmap
- [ ] Split the code into modules (core logic / UI)
- [ ] REST API (FastAPI) so a mobile app can use the same logic
- [ ] User login and per-user quota
- [ ] Save papers and a question bank
- [ ] Auto-grade MCQs in code from the answer key

---
Designed by Waheed Ali Hamouzai - WSA Educational Community
