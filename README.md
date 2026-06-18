# recruiter-agent

A recruitment assistant agent. It turns uploaded CVs (PDF/DOCX) into rows in an in-memory
candidate database, flags duplicates/cross-role/cross-TA conflicts, and syncs a personal
database to a shared team database — all through a chat UI (`index.html`).

Runs entirely on GreenNode AgentBase (`main.py`) — there is no separate local-dev server and
no local Excel/Windows-path dependency; `main.py` is the same entrypoint for local runs and the
cloud image.

## Prerequisites

- Python 3.10+
- A GreenNode IAM Service Account ([create one here](https://iam.console.vngcloud.vn/service-accounts))

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
   - `TA_INCHARGE` — your name as it appears in the "TA Incharge" column (defaults to Windows login)

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

`main.py` imports the `greennode_agentbase` SDK, which prints a `✓` via `rich` at startup — on
Windows terminals using the legacy `cp1252` codepage this crashes on import. Force UTF-8 first:

```bash
# Windows (PowerShell):
$env:PYTHONUTF8 = "1"; python main.py

# macOS/Linux/Git Bash:
PYTHONIOENCODING=utf-8 PYTHONUTF8=1 python main.py
```

It starts on `http://127.0.0.1:8080` (bound to `0.0.0.0:8080`).

Open `index.html` directly in a browser (double-click the file, or `file://` it).

From the chat UI you can create job folders, upload CVs, resolve duplicate/conflict alerts via
buttons, and trigger a database sync.

### Endpoints

| Method | Path                     | Purpose                                          |
|--------|--------------------------|---------------------------------------------------|
| POST   | `/invocations`           | Chat with the agent: `{"message": "..."}`         |
| POST   | `/save-cv`                | Stage an uploaded CV in memory (no AI yet), multipart: `folder`, `subfolder`, `file`, `referrer` |
| POST   | `/start-parse`            | Start background AI parsing of staged CVs: `{"files": [...]}` |
| GET    | `/parse-status/{job_id}` | Poll a background parse job's status              |
| GET    | `/list-folders`           | List known job folders                            |
| GET    | `/database`               | `{"personal": [...], "team": [...]}` — both databases as JSON |
| GET    | `/download-excel`         | Download the personal database as `.xlsx`         |
| GET    | `/health`                 | `{"status": "Healthy"}`                           |

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

The Docker image runs `main.py` (`Dockerfile` `CMD`) — the same entrypoint used for local dev,
wrapped in the `GreenNodeAgentBaseApp` SDK. All endpoints above (`/invocations`, `/save-cv`,
`/start-parse`, `/parse-status`, `/list-folders`, `/database`, `/download-excel`, `/health`) are
available on the deployed cloud endpoint too.

1. Build and push your Docker image (or use `/agentbase-deploy` skill)
2. Create a Runtime at https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime
3. Create an Endpoint pointing to your Runtime

See the [AgentBase Console](https://aiplatform.console.vngcloud.vn) to manage runtimes, identities, and memory.

## Storage

Candidate data (personal DB, team DB, job folders, alerts, sync state) lives in module-level
lists inside `main.py`, write-through to JSON files under `data/` for crash safety within a
single running container. There is no generic AgentBase persistent-storage service — **a
redeploy or container restart resets all data.** Avoid redeploying between seeding data and
using it for a demo.

## Add Conversation Memory (Optional)

When you need conversation history or long-term memory, use `/agentbase-memory` to set up AgentBase Memory and integrate it with your agent.

## Project Structure

- `main.py` - Single entrypoint (`GreenNodeAgentBaseApp`), used for both local runs and the Docker image
- `index.html` - Chat UI: job folders, CV upload, duplicate/conflict resolution, sync controls
- `Dockerfile` - Container image definition (runs `main.py`)
- `requirements.txt` - Python dependencies
- `.greennode.json` - AgentBase configuration (cloud deployment)
- `.env.example` - Environment variable template
- `data/` - runtime state (personal/team DB, job folders, alerts, sync snapshot) — gitignored, not durable across redeploys (see [Storage](#storage))
