from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openai import OpenAI
import os
import json
from pathlib import Path

app = FastAPI(title="AMPM Copilot API")

allowed_origins = os.getenv(
    "ALLOWED_ORIGINS",
    "http://127.0.0.1:5500,http://localhost:5500,https://ampmsal.github.io"
).split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in allowed_origins if o.strip()],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATA_DIR = Path("/app/data_api/out")
MODEL = os.getenv("OPENAI_MODEL", "gpt-5.4-mini")
MAX_OUTPUT_TOKENS = int(os.getenv("OPENAI_MAX_OUTPUT_TOKENS", "700"))

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

class QueryRequest(BaseModel):
    question: str
    page_context: dict = {}
    filters: dict = {}

def load_catalog():
    files = []
    if DATA_DIR.exists():
        for p in DATA_DIR.rglob("*.json"):
            files.append(str(p.relative_to(DATA_DIR)).replace("\\", "/"))
    return sorted(files)

@app.get("/health")
def health():
    return {"ok": True, "files": load_catalog()}

@app.get("/api/catalog")
def catalog():
    return {
        "ok": True,
        "files": load_catalog(),
        "message": "Backend activo y leyendo JSON"
    }

@app.post("/api/query")
def query(req: QueryRequest):
    files = load_catalog()

    facts = {
        "question": req.question,
        "page_context": req.page_context,
        "filters": req.filters,
        "available_json_files": files[:80]
    }

    prompt = f"""
Eres un copiloto analítico de retail.
No inventes datos.
No digas que analizaste archivos específicos si no aparecen en available_json_files.
Tu tarea es devolver SOLO JSON válido con esta estructura:

{{
  "title": "string",
  "summary": "string",
  "actions": ["string", "string"],
  "chart": {{
    "type": "bar|line|pie|doughnut|none",
    "title": "string",
    "labels": [],
    "datasets": [{{"label":"string","data":[]}}]
  }},
  "notes": ["string"]
}}

Contexto:
{json.dumps(facts, ensure_ascii=False)}
"""

    response = client.responses.create(
        model=MODEL,
        input=prompt,
        max_output_tokens=MAX_OUTPUT_TOKENS
    )

    text = response.output_text

    try:
        parsed = json.loads(text)
        return {"ok": True, "result": parsed}
    except Exception:
        return {
            "ok": True,
            "result": {
                "title": "Respuesta generada",
                "summary": text,
                "actions": [],
                "chart": {"type": "none", "title": "", "labels": [], "datasets": []},
                "notes": ["La respuesta no salió en JSON perfecto todavía."]
            }
        }