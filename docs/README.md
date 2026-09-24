# DQE User Guide

`dqe-user-guide.pdf` is a 2-page, branded user guide for the DQE app — how to get it running locally, and what the main features do.

## Regenerating it

The PDF is generated from `generate_user_guide.py`, not hand-edited. To update it (e.g. after a UI/branding change or a new feature worth documenting):

```bash
pip3 install -r docs/requirements.txt   # one-time, only needed for doc generation
python3 docs/generate_user_guide.py
```

This writes `docs/dqe-user-guide.pdf` in place. `logo_white.png` is a cached, auto-generated asset (a white version of `static/duetto-logo.png`, matching the CSS filter the app itself uses on the dark nav bar) — it regenerates automatically if the source logo changes, no need to touch it manually.
