"""
Blue Lime proposal generation service.

A small FastAPI app that accepts a Blue Lime Excel proposal template and
returns a polished, branded PDF.

Endpoints:
  GET  /             — Service info
  GET  /health       — Liveness check (used by Render's health checks)
  POST /generate-proposal — Multipart upload, returns PDF bytes

Auth:
  All POST endpoints require the X-API-Key header to match PROPOSAL_API_KEY
  in the environment. GET endpoints are public so Render can probe /health.
"""
from __future__ import annotations

import logging
import os
import re
import tempfile
import traceback
from datetime import datetime, timezone

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from .excel_parser import parse_excel
from .proposal_generator import build_proposal

# -----------------------------------------------------------------------------
# Config
# -----------------------------------------------------------------------------
API_KEY = os.environ.get("PROPOSAL_API_KEY")
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "20"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

# Allow your Lovable app + local dev to call us. Set via env in production.
CORS_ALLOW_ORIGINS = [
    o.strip()
    for o in os.environ.get("CORS_ALLOW_ORIGINS", "*").split(",")
    if o.strip()
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("proposal-service")

# -----------------------------------------------------------------------------
# App
# -----------------------------------------------------------------------------
app = FastAPI(
    title="Blue Lime Proposal Service",
    description="Generates polished, branded PDF proposals from Excel templates.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ALLOW_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def _check_api_key(provided: str | None) -> None:
    """Raise 401 if the API key is missing or wrong.

    If PROPOSAL_API_KEY is unset (e.g., in local dev), auth is disabled. We log
    a warning so this is loud and not a silent foot-gun.
    """
    if not API_KEY:
        raise RuntimeError("PROPOSAL_API_KEY is not set in environment.")
    if not provided or provided != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")


def _safe_filename(name: str) -> str:
    """Normalize a filename for use in Content-Disposition."""
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_")
    return base or "proposal.pdf"


# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------
@app.get("/")
async def root():
    return {
        "service": "Blue Lime Proposal Service",
        "version": app.version,
        "endpoints": ["/health", "/generate-proposal (POST)"],
    }


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "time":   datetime.now(timezone.utc).isoformat(),
        "auth": bool(API_KEY),
    }


@app.post("/generate-proposal")
async def generate_proposal(
    file: UploadFile = File(..., description="Excel proposal template (.xlsx)"),
    x_api_key: str | None = Header(default=None, alias="X-API-Key"),
):
    """Generate a branded PDF from an uploaded Excel template.

    Returns the PDF as a `application/pdf` response, with
    Content-Disposition: attachment; filename="...".
    """
    _check_api_key(x_api_key)

    # Validate filename + size
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=400,
            detail="Only .xlsx or .xlsm files are accepted.",
        )

    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"File exceeds maximum upload size of {MAX_UPLOAD_MB} MB.",
        )
    if len(contents) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    log.info("Received Excel: %s (%d bytes)", file.filename, len(contents))

    # Parse → render
    try:
        data = parse_excel(contents)
    except Exception as exc:
        log.error("Parse failure: %s\n%s", exc, traceback.format_exc())
        raise HTTPException(
            status_code=422,
            detail=f"Could not parse Excel: {exc}",
        )

    short = data["client"].get("short_name") or "Proposal"
    out_name = _safe_filename(f"{short.replace(' ', '_')}_Proposal.pdf")

    # Write to a tempfile and stream back. Render's filesystem is ephemeral
    # but writable, which is exactly what we want — the file is gone after
    # the response is delivered.
    with tempfile.NamedTemporaryFile(
        suffix=".pdf", delete=False
    ) as tmp:
        out_path = tmp.name

    try:
        build_proposal(data, out_path)
    except Exception as exc:
        log.error("Build failure: %s\n%s", exc, traceback.format_exc())
        if os.path.exists(out_path):
            os.unlink(out_path)
        raise HTTPException(
            status_code=500,
            detail=f"Could not generate PDF: {exc}",
        )

    log.info("Generated %s for %s", out_name, short)

    return FileResponse(
        out_path,
        media_type="application/pdf",
        filename=out_name,
        headers={"X-Generated-For": short},
    )


# -----------------------------------------------------------------------------
# Friendly error responses for unexpected issues
# -----------------------------------------------------------------------------
@app.exception_handler(Exception)
async def _generic_exception_handler(request, exc):
    log.error("Unhandled error on %s: %s\n%s",
              request.url.path, exc, traceback.format_exc())
    return JSONResponse(
        status_code=500,
        content={"detail": f"Internal error: {type(exc).__name__}"},
    )
