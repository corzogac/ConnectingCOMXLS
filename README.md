# Connecting Excel (COM) + VBA Tank Model — UV workflow

Run an Excel **VBA tank model** from Python via COM, using a clean **uv** environment.  
Data flow: Python writes precipitation to `input` sheet → VBA model runs → results written to `discharge` → Python reads (or you view in Excel).

## Requirements
- Windows + Microsoft **Excel (desktop)**
- Python 3.11+ (uv will manage the venv)
- [uv](https://docs.astral.sh/uv/) installed

## Repo structure
- `README.md` (this file)
- `requirements.txt`
- `uv.toml`
- `run.py`