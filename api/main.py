import os
import time
from io import BytesIO
from urllib.parse import urlparse

import httpx
import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook

app = FastAPI(title="AMPM Data API", version="1.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

CACHE = {
    "ts": 0,
    "excel_bytes": None,
    "sheet_names": None,
}

CACHE_SECONDS = 300


async def ensure_excel_bytes():
    excel_url = os.getenv("EXCEL_URL", "").strip()
    if not excel_url:
        raise HTTPException(status_code=500, detail="Falta la variable EXCEL_URL")

    now = time.time()
    if CACHE["excel_bytes"] is not None and (now - CACHE["ts"]) < CACHE_SECONDS:
        return CACHE["excel_bytes"]

    async with httpx.AsyncClient(timeout=120, follow_redirects=True) as client:
        response = await client.get(excel_url)
        response.raise_for_status()
        content = response.content

    CACHE["ts"] = now
    CACHE["excel_bytes"] = content
    CACHE["sheet_names"] = None
    return content


@app.get("/health")
async def health():
    return {"ok": True}


@app.get("/env-check")
async def env_check():
    excel_url = os.getenv("EXCEL_URL", "").strip()
    parsed = urlparse(excel_url) if excel_url else None
    return {
        "ok": True,
        "excel_url_set": bool(excel_url),
        "excel_url_host": parsed.netloc if parsed else None,
        "excel_url_path_tail": parsed.path[-80:] if parsed else None,
    }


@app.get("/download-check")
async def download_check():
    content = await ensure_excel_bytes()
    return {
        "ok": True,
        "bytes": len(content)
    }


@app.get("/meta")
async def meta():
    if CACHE["sheet_names"] is not None:
        return {
            "ok": True,
            "sheets": CACHE["sheet_names"]
        }

    content = await ensure_excel_bytes()

    try:
        wb = load_workbook(filename=BytesIO(content), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No pude leer nombres de hojas: {str(e)}")

    CACHE["sheet_names"] = sheet_names

    return {
        "ok": True,
        "sheets": sheet_names
    }


@app.get("/sheet-preview")
async def sheet_preview(
    sheet: str = Query(...),
    rows: int = Query(5, ge=1, le=20)
):
    content = await ensure_excel_bytes()

    try:
        df = pd.read_excel(BytesIO(content), sheet_name=sheet, nrows=rows)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No pude leer la hoja {sheet}: {str(e)}")

    df.columns = [str(c).strip() for c in df.columns]

    return {
        "ok": True,
        "sheet": sheet,
        "columns": df.columns.tolist(),
        "preview_rows": df.fillna("").to_dict(orient="records")
    }