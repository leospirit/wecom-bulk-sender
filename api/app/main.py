from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles

from .db import init_db
from .routes import router

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"] ,
    allow_headers=["*"],
)

@app.on_event("startup")
def _startup():
    init_db()

app.include_router(router, prefix="/api")

_static_root = Path(__file__).resolve().parent / "static" / "rpa-lite"
if _static_root.exists():
    app.mount("/rpa-lite", StaticFiles(directory=str(_static_root), html=True), name="rpa-lite")


@app.get("/")
def root_redirect():
    if _static_root.exists():
        return RedirectResponse(url="/rpa-lite/")
    return {"status": "ok"}
