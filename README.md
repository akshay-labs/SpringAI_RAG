# HR Policy RAG Demo

A small end-to-end RAG (Retrieval-Augmented Generation) application built for a Day 2/3 training
session:

- **Day 2** — vector embeddings, chunking, MariaDB vector search
- **Day 3** — full RAG pipeline: ingestion → embeddings → retrieval → LLM answer, exposed via a
  Spring Boot REST API with a minimal React front end

It ingests two generated, fictional policy PDFs — an **HR Company Policy** and an **IT & Data
Security Policy** — splits them into chunks, embeds them into MariaDB's native vector store, and
answers questions grounded in whichever policy is relevant. Every question is also answered by the
plain LLM with no retrieval at all, side by side, so the effect of grounding is visible directly.

## Architecture

```
PDF (bundled)  →  PagePdfDocumentReader  →  TokenTextSplitter  →  chunks
                                                                     │
                                                    OpenAI embeddings (text-embedding-3-small)
                                                                     │
                                                                     ▼
                                                     MariaDB vector_store table
                                                                     ▲
question  →  similaritySearch(topK) ──────────────────────────────┘
   │
   ▼
augmented prompt (context + question)  →  OpenAI chat (gpt-4o-mini)  →  answer
```

Both chat and embeddings use **OpenAI** - one provider, one API key, no local model runner. (An
earlier version of this project tried local Ollama for chat and a local ONNX model for
embeddings; both ran into machine-specific friction - Ollama's executable being blocked by
corporate endpoint security, and native-library extraction issues with the ONNX runtime - so this
settled on the simplest path: a single cloud provider for both.)

## Prerequisites

- Java 17+ (tested with the JDK at `C:\Program Files\Java\jdk-17`)
- Maven (tested with 3.9.9)
- Node.js 18+ / npm (tested with Node v25.8.0 / npm 11.11.0)
- MariaDB 11.7+ (tested with 12.1) running locally
- An OpenAI API key from platform.openai.com, with billing/credits set up

## One-time setup

**1. Create the database:**

```bash
mysql -u root -p -e "CREATE DATABASE ragdb;"
```

(If your MariaDB root user has no password, drop `-p`. Update `DB_USERNAME` / `DB_PASSWORD` env
vars if you use a different account/port.)

**2. Set your OpenAI API key:**

```bash
set OPENAI_API_KEY=sk-...
```

**3. (Optional) Regenerate the seed PDFs** — only needed if you want to edit the policy text:

```bash
pip install reportlab
python tools/generate_hr_policy_pdf.py
python tools/generate_it_security_policy_pdf.py
```

## Running the backend

```bash
mvn spring-boot:run
```

The app starts on `http://localhost:8080` and auto-creates the `vector_store` table in `ragdb` on
first run (`spring.ai.vectorstore.mariadb.initialize-schema=true`).

### API

| Method | Path                      | Body                          | Description                                   |
|--------|---------------------------|--------------------------------|------------------------------------------------|
| POST   | `/api/v1/documents/ingest`| —                               | Reads both bundled PDFs, chunks them, embeds + stores them |
| GET    | `/api/v1/documents/status`| —                               | Returns how many chunks are currently stored  |
| POST   | `/api/v1/rag/query`       | `{"question": "..."}`           | Runs retrieval + LLM answer, returns `{answer, sources[]}` |
| POST   | `/api/v1/rag/compare`     | `{"question": "..."}`           | Answers once with no retrieval and once with RAG, returns `{question, directAnswer, ragAnswer, sources[]}` — `sources[].score` is the cosine similarity for that chunk |

Example:

```bash
curl -X POST http://localhost:8080/api/v1/documents/ingest

curl -X POST http://localhost:8080/api/v1/rag/compare \
  -H "Content-Type: application/json" \
  -d "{\"question\":\"How many days of annual leave do employees get?\"}"
```

Try an HR question and an IT question (e.g. "What should I do if I lose my company laptop?") back
to back — the `sources[]` in each response come from the matching document, even though both
policies live in the same `vector_store` table.

## Running the frontend

```bash
cd rag-demo-ui
npm install
npm run dev
```

Open the printed local URL (default `http://localhost:5173`). Click **Ingest Both Policy PDFs**
once, then ask a question — the page shows the plain-LLM answer and the RAG answer side by side,
with the matched source chunks and their similarity scores underneath the RAG answer.

## Project layout

```
pom.xml                              Backend Maven build
src/main/java/com/training/ragdemo/
  RagDemoApplication.java
  config/CorsConfig.java             Allows the Vite dev server to call the API
  service/DocumentIngestionService.java   PDF → chunks → vector store
  service/RagQueryService.java            similarity search → augmented prompt → LLM
  web/RagController.java             REST endpoints
  web/dto/                           Request/response records
src/main/resources/
  application.yml                    Config (OpenAI chat + embeddings, MariaDB)
  documents/HR-Company-Policy.pdf     Seed document #1
  documents/IT-Security-Policy.pdf    Seed document #2
tools/generate_hr_policy_pdf.py             Regenerates seed PDF #1
tools/generate_it_security_policy_pdf.py    Regenerates seed PDF #2
rag-demo-ui/                          Vite + React front end
```

## Possible next steps (not built yet)

- Upload endpoint to ingest an arbitrary PDF instead of the bundled ones
- Streaming answers (SSE) instead of a blocking response
- A second embedding/vector-store backend (e.g. Chroma) to compare against MariaDB
