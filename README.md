# Excel Filler Service

FastAPI service that fills the Procapita Excel template from Supabase data.
Deploy on Railway in 2 minutes.

## Deploy to Railway

1. Push this folder to a GitHub repo (or use Railway CLI)
2. In Railway: **New Project → Deploy from GitHub repo**
3. Set environment variables (Railway → Service → Variables):
   - `SUPABASE_URL` — your Supabase project URL
   - `SUPABASE_KEY` — your Supabase **service role** key (not anon key)
4. Railway auto-detects Python and deploys. Your service URL will be something like `https://excel-filler-service.up.railway.app`

## API

### `POST /fill`

Fills the Excel template and uploads the result to Supabase Storage.

**Request body:**
```json
{
  "excel_base64": "<base64 of the .xlsx file>",
  "excel_filename": "n8nworking.xlsx",
  "target_year": 2026,
  "bucket": "pipeline-uploads",
  "run_id": "20260628-120000"
}
```

**Response:**
```json
{
  "ok": true,
  "total_written": 5686,
  "unmatched_count": 1,
  "download_url": "https://your-project.supabase.co/storage/v1/...",
  "filename": "filled-20260628-120000-n8nworking.xlsx",
  "target_year": 2026
}
```

If Supabase Storage upload fails, `download_url` will be `null` and `excel_base64` will contain the filled file as base64.

### `GET /health`
Returns `{"status": "ok"}`.

## n8n Integration

Add an **HTTP Request** node after `Run Forecast (02b)` in the 01 Orchestrator:

- **Method:** POST
- **URL:** `https://your-service.up.railway.app/fill`
- **Body (JSON):**
```json
{
  "excel_base64": "{{ $json.excel_base64 }}",
  "excel_filename": "{{ $json.excel_filename }}",
  "target_year": {{ $json.target_year }},
  "bucket": "{{ $json.bucket }}",
  "run_id": "{{ $json.run_id }}"
}
```
