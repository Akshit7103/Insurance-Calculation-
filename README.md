# MB Calculator Web App

Minimal FastAPI web app for calculating the `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP` output columns.

## Local Run

```powershell
python -m pip install -r requirements.txt
python -m uvicorn web_app:app --host 127.0.0.1 --port 8000
```

Open:

```text
http://127.0.0.1:8000/
```

On Windows, you can also double-click:

```text
launch.bat
```

## Render Deployment

Push this folder to a GitHub repository, then create a new Render Web Service from that repo.

Use these settings if Render does not auto-detect them from `render.yaml`:

```text
Environment: Python
Build Command: pip install -r requirements.txt
Start Command: uvicorn web_app:app --host 0.0.0.0 --port $PORT
Health Check Path: /health
```

## Input and Output

Input:

- One Excel workbook or CSV file.
- One sheet, preferably named `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP`.
- Client data columns `A:BD`.

Output:

- One Excel workbook.
- One sheet named `MBP_CPP_CHP_RFP_SMB_MSB_RMM_MMP`.
- Input columns `A:BD` plus calculated columns `BF:DO`.
- Calculated columns `BF:DF` are hidden in the generated workbook.

The generated Excel output contains calculated values. The bonus-sheet logic used by `BZ`, `CH`, `CO`, and `CS` is implemented in `mb_calculator.py`, so the uploaded workbook does not need separate bonus sheets.
