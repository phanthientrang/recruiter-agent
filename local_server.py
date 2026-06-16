"""
Local HTTP server for recruiter-agent.

Usage:
  python local_server.py
  ngrok http 8080          (in a second terminal)
  → paste the ngrok URL into the Settings bar in the UI

Endpoints:
  POST /invocations    {"message": "..."}                → agent response (JSON)
  POST /upload-cv      multipart: folder, subfolder, file → saves file + agent response (JSON)
  GET  /health                                            → {"status": "healthy"}

main.py is imported for its compiled graph and helpers only.
main.py's app.run() is guarded by `if __name__ == "__main__"`, so no
AgentBase cloud server starts when this module imports it.
"""

import asyncio
import os
from datetime import datetime
from pathlib import Path

import uvicorn
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.cors import CORSMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

# Import the compiled LangGraph and helpers from main.py.
# Side-effects on import:
#   - load_dotenv() runs (reads .env)
#   - LLM client is created (no network call yet)
#   - Graph is compiled (pure in-memory)
#   - GreenNodeAgentBaseApp is instantiated (sets up ASGI app, no cloud calls)
# app.run() is NOT called because it is inside `if __name__ == "__main__"`.
from main import graph, _process_cv, JOBS_BASE_DIR, JOB_SUB_FOLDERS  # noqa: E402

_ALLOWED_EXT = {".pdf", ".docx"}


# ---------------------------------------------------------------------------
# POST /invocations
# ---------------------------------------------------------------------------
async def handle_invocations(request: Request) -> JSONResponse:
    try:
        body = await request.json()
    except Exception:
        return JSONResponse(
            {"status": "error", "response": "Request body must be valid JSON."},
            status_code=400,
        )

    message = str(body.get("message", "")).strip()
    if not message:
        return JSONResponse(
            {"status": "error", "response": "Field 'message' is required."},
            status_code=400,
        )

    try:
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None,
            lambda: graph.invoke({"messages": [("user", message)]}),
        )
        return JSONResponse({
            "status": "success",
            "response": result["messages"][-1].content,
            "timestamp": datetime.now().isoformat(),
        })
    except Exception as exc:
        return JSONResponse(
            {
                "status": "error",
                "response": f"Agent error: {exc}",
                "timestamp": datetime.now().isoformat(),
            },
            status_code=500,
        )


# ---------------------------------------------------------------------------
# POST /upload-cv   (multipart/form-data: folder, subfolder, file)
# ---------------------------------------------------------------------------
async def handle_upload_cv(request: Request) -> JSONResponse:
    try:
        form = await request.form()
    except Exception as exc:
        return JSONResponse(
            {
                "status": "error",
                "response": (
                    f"Could not parse multipart form data ({exc}). "
                    "Make sure python-multipart is installed: pip install python-multipart"
                ),
            },
            status_code=400,
        )

    folder    = str(form.get("folder")    or "").strip()
    subfolder = str(form.get("subfolder") or "").strip()
    upload    = form.get("file")

    if not folder or not subfolder or upload is None:
        return JSONResponse(
            {"status": "error", "response": "Fields 'folder', 'subfolder', and 'file' are all required."},
            status_code=400,
        )

    if subfolder not in JOB_SUB_FOLDERS:
        return JSONResponse(
            {
                "status": "error",
                "response": f"'subfolder' must be one of: {', '.join(JOB_SUB_FOLDERS)}",
            },
            status_code=400,
        )

    filename = Path(upload.filename).name  # strip path traversal
    ext = Path(filename).suffix.lower()
    if ext not in _ALLOWED_EXT:
        return JSONResponse(
            {
                "status": "error",
                "response": f"File type '{ext}' is not supported. Use PDF or DOCX.",
            },
            status_code=400,
        )

    # Save file to: JOBS_BASE_DIR / [folder] / [subfolder] / [filename]
    save_dir  = Path(JOBS_BASE_DIR) / folder / subfolder
    save_dir.mkdir(parents=True, exist_ok=True)
    save_path = save_dir / filename

    contents = await upload.read()
    with open(save_path, "wb") as fh:
        fh.write(contents)

    # Ask the agent to process the saved file
    message = f"Process the CV file at: {save_path}"
    try:
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None,
            lambda: graph.invoke({"messages": [("user", message)]}),
        )
        return JSONResponse({
            "status": "success",
            "response": result["messages"][-1].content,
            "saved_to": str(save_path),
            "timestamp": datetime.now().isoformat(),
        })
    except Exception as exc:
        return JSONResponse(
            {
                "status": "error",
                "response": (
                    f"CV saved to:\n{save_path}\n\n"
                    f"But agent processing failed: {exc}"
                ),
                "saved_to": str(save_path),
                "timestamp": datetime.now().isoformat(),
            },
            status_code=500,
        )


# ---------------------------------------------------------------------------
# GET /health
# ---------------------------------------------------------------------------
async def handle_health(request: Request) -> JSONResponse:
    return JSONResponse({
        "status": "healthy",
        "jobs_dir": JOBS_BASE_DIR,
        "timestamp": datetime.now().isoformat(),
    })


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
app = Starlette(
    routes=[
        Route("/invocations", handle_invocations, methods=["POST"]),
        Route("/upload-cv",   handle_upload_cv,   methods=["POST"]),
        Route("/health",      handle_health,       methods=["GET"]),
    ],
    middleware=[
        Middleware(
            CORSMiddleware,
            allow_origins=["*"],
            allow_methods=["GET", "POST", "OPTIONS"],
            allow_headers=["*"],
        )
    ],
)

if __name__ == "__main__":
    print()
    print("  ┌─────────────────────────────────────────────────┐")
    print("  │  Recruiter Agent — Local Server                 │")
    print(f"  │  Jobs dir : {JOBS_BASE_DIR[:38]:<38} │")
    print("  │  URL      : http://0.0.0.0:8080                 │")
    print("  │  Tunnel   : ngrok http 8080                     │")
    print("  └─────────────────────────────────────────────────┘")
    print()
    uvicorn.run(app, host="0.0.0.0", port=8080, log_level="info")
