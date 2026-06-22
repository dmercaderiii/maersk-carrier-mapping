# Maersk Quote Automation

Local web automation for uploading a Maersk quote workbook, flattening the `AFLS Quote` rows into a new `Processed Output` sheet, and downloading the processed `.xlsx` file automatically.

## What it does

- Reads the `AFLS Quote` sheet from the uploaded workbook.
- Treats each `BAS` row as the main lane row.
- Keeps columns `A:P` from the `BAS` row on a single output line.
- Converts the charges below `BAS` into new columns starting after `P`.
- Uses the format `<CHARGE>_<CONTAINER>_<RATE_BASIS>_CURRENCY` for the generated headers.
- Returns the workbook with a new `Processed Output` sheet added.

## Run locally

```powershell
python server.py
```

Then open [http://127.0.0.1:8000](http://127.0.0.1:8000).

## Deploy on Render

This app is ready for Render.

1. Push the latest code to GitHub.
2. In Render, create a new `Web Service`.
3. Connect the repository `dmercaderiii/maersk-carrier-mapping`.
4. Render can use the included [render.yaml](C:/Users/Dimidim/OneDrive%20-%20Freight%20Right%20Global%20Logistics/Desktop/Automation%20Projects/Maersk%20Automation/render.yaml), or you can enter these settings manually:
   - Runtime: `Python`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python server.py`
5. Deploy.

The server automatically binds to Render's `PORT` environment variable.

## Notes

- Input format currently supports `.xlsx`.
- The processor looks for a sheet named `AFLS Quote`, then falls back to any sheet containing `Quote` in its name.
- If a `Processed Output` sheet already exists, it is replaced on each run.
