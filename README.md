# recruiter-agent

A recruitment assistant agent. It turns CVs (PDF/DOCX) dropped into job folders into rows in an
Excel candidate database, flags duplicates/cross-role/cross-TA conflicts, and syncs a personal
database to a shared team database — all through a chat UI (`index.html`).

Built on GreenNode AgentBase for cloud deployment; for local development it runs as a
standalone Starlette server (`local_server.py`).

## Prerequisites

- Python 3.10+
- A GreenNode IAM Service Account — only needed for [cloud deployment](#deploy-to-agentbase-runtime), not local dev ([create one here](https://iam.console.vngcloud.vn/service-accounts))
- [ngrok](https://ngrok.com/) — only needed if you want to open the chat UI from another device

## Setup

1. Create and activate a virtual environment:
   ```bash
   # macOS/Linux:
   python3 -m venv venv && source venv/bin/activate

   # Windows (PowerShell):
   python -m venv venv; venv\Scripts\Activate.ps1
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure environment variables:
   ```bash
   cp .env.example .env
   ```
   Edit `.env` and fill in:
   - `LLM_API_KEY`, `LLM_BASE_URL`, `LLM_MODEL` — see [Configure LLM](#configure-llm)
   - `JOBS_BASE_DIR` — root folder containing all job folders
   - `EXCEL_PATH` — path to your personal candidate database (`.xlsx`/`.xlsm`)
   - `TA_INCHARGE` — your name as it appears in the "TA Incharge" column (defaults to Windows login)
   - `TEAM_EXCEL_PATH` — path to the shared team database (used by sync)
   - `CV_POLL_INTERVAL` — seconds between background CV folder scans (default `60`)

   `GREENNODE_CLIENT_ID` / `GREENNODE_CLIENT_SECRET` / `GREENNODE_AGENT_IDENTITY` (or `.greennode.json`)
   are only needed for [cloud deployment](#deploy-to-agentbase-runtime) — on AgentBase Runtime these
   are managed automatically and don't need to be set locally.

## Configure LLM

This project uses any OpenAI-compatible LLM provider. Set the following in `.env`:

```
LLM_API_KEY=your-api-key
LLM_BASE_URL=your-provider-base-url
LLM_MODEL=your-model-name
```

**Provider examples:**
- **GreenNode AIP**: Use `/agentbase-llm` to get an API key. Set `LLM_BASE_URL=https://maas-llm-aiplatform-hcm.api.vngcloud.vn/v1`
- **OpenAI**: Set `LLM_BASE_URL=https://api.openai.com/v1`, model e.g. `gpt-4o`
- **Ollama** (local): Set `LLM_BASE_URL=http://localhost:11434/v1` (no key needed)

**Production**: Use `/agentbase-identity` to store your API key on the platform and inject it at runtime.

## Run Locally

Local dev uses `local_server.py`, **not** `main.py`. `main.py` imports the `greennode_agentbase`
SDK, which crashes on import in Windows `cp1252` terminals (it prints a `✓` via `rich`); `main.py`
is kept only for [cloud deployment](#deploy-to-agentbase-runtime).

1. Start the server:
   ```bash
   python local_server.py
   ```
   It starts on `http://127.0.0.1:8080` (bound to `0.0.0.0:8080`).

2. Open `index.html` directly in a browser (double-click the file, or `file://` it).

3. Click the ⚙️ gear icon in the UI and set the server URL:
   - Same machine: `http://127.0.0.1:8080`
   - Different device / phone: run `ngrok http 8080`, then paste the `https://xxx.ngrok-free.app` URL

From the chat UI you can create job folders, upload CVs, resolve duplicate/conflict alerts via
buttons, and trigger a database sync.

### Endpoints

| Method | Path                     | Purpose                                          |
|--------|--------------------------|---------------------------------------------------|
| POST   | `/invocations`           | Chat with the agent: `{"message": "..."}`         |
| POST   | `/save-cv`                | Save an uploaded CV to disk (no AI), multipart: `folder`, `subfolder`, `file` |
| POST   | `/start-parse`            | Start background AI parsing of saved CVs: `{"files": [...]}` |
| GET    | `/parse-status/{job_id}` | Poll a background parse job's status              |
| GET    | `/list-folders`           | List job folders under `JOBS_BASE_DIR`            |
| GET    | `/open-folder?path=...`  | Open a job folder in Windows Explorer             |
| GET    | `/health`                 | `{"status": "healthy"}`                           |

Test the chat endpoint directly:
```bash
curl -X POST http://127.0.0.1:8080/invocations \
  -H "Content-Type: application/json" \
  -d '{"message": "Hello, agent!"}'
```

Health check:
```bash
curl http://127.0.0.1:8080/health
```

## Deploy to AgentBase Runtime

The Docker image runs `main.py` (`Dockerfile` `CMD`), which wraps the same agent logic in the
`GreenNodeAgentBaseApp` SDK and only exposes `/invocations` + a health check — the `local_server.py`
file-upload/sync endpoints are local-dev only and are not part of the cloud image.

1. Build and push your Docker image (or use `/agentbase-deploy` skill)
2. Create a Runtime at https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime
3. Create an Endpoint pointing to your Runtime

See the [AgentBase Console](https://aiplatform.console.vngcloud.vn) to manage runtimes, identities, and memory.

## Add Conversation Memory (Optional)

When you need conversation history or long-term memory, use `/agentbase-memory` to set up AgentBase Memory and integrate it with your agent.

## Project Structure

- `main.py` - Cloud entrypoint (`GreenNodeAgentBaseApp`), used by the Docker image only
- `local_server.py` - Local dev server (Starlette/uvicorn) with the CV-upload/parse/sync endpoints
- `index.html` - Chat UI: CV upload, duplicate/conflict resolution, sync controls
- `Dockerfile` - Container image definition (runs `main.py`)
- `requirements.txt` - Python dependencies
- `.greennode.json` - AgentBase configuration (cloud deployment)
- `.env.example` - Environment variable template
- `processed_cvs.json`, `sync_state.json` - local runtime state (which CVs were processed, last sync snapshot)
- `data/alerts.json` - pending alerts (duplicates, conflicts, sync results) shown in the chat UI
