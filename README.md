# Excel AI Assistant — Office.js Add-in

An AI-powered Excel add-in that lets users chat with their spreadsheet. The user types a natural-language request; an AI agent figures out the required Excel operations and executes them directly inside the workbook.

---

## Stack

| Layer | Tech |
|---|---|
| Add-in frontend | React + TypeScript + Vite (Office.js) |
| AI agent loop | Custom agentic loop in `src/services/agent.ts` |
| Excel operations | Office.js API via `src/services/office.ts` |
| Backend / LLM proxy | Node.js + Express (`backend/index.js`) |
| LLM | OpenAI GPT (via `openai` SDK) |
| Auth | Auth0 (JWT — RS256, validated on every backend request) |

---

## How It Works (Step by Step)

1. **User logs in** via Auth0. The add-in gets a JWT access token.

2. **User types a request** in the chat (e.g. *"Sort by sales column descending"*).

3. **Agent reads the workbook context** — active sheet, all sheet metadata, selected range values, and a 50-row data preview — using `getWorkbookContext()` (Office.js).

4. **Agent calls the backend** (`POST /query-answer-final-output`) with the full context + user request. The JWT is sent in the `Authorization: Bearer` header.

5. **Backend validates the JWT** against Auth0's JWKS endpoint, then forwards the prompt to the OpenAI API.

6. **LLM responds with tool calls** as raw JSON objects (e.g. `{"tool":"sort_range","sheet":"Sheet1","range":"A2:C10","sortColumn":1,"ascending":false}`). It can return a single call or a batch array.

7. **Agent parses and executes each tool call** via `executeTool()` (Office.js), collecting results into a history log.

8. **If more steps are needed**, the agent re-calls the backend with the history and asks for the next set of tool calls (max 15 iterations).

9. **When the LLM returns `final_answer`**, the agent stops and shows the summary to the user in the chat.

---

## Running Locally

```bash
# Frontend (add-in dev server)
npm install
npm run dev          # starts on https://localhost:3000

# Backend
cd backend
npm install
node index.js        # starts on http://localhost:8000
```

**Required `.env` files:**

`backend/.env`
```
AUTH0_DOMAIN=your-tenant.auth0.com
AUTH0_AUDIENCE=your-api-identifier
OPENAI_API_KEY=sk-...
PORT=8000
```

`.env` (frontend)
```
VITE_AUTH0_DOMAIN=your-tenant.auth0.com
VITE_AUTH0_CLIENT_ID=your-client-id
VITE_AUTH0_AUDIENCE=your-api-identifier
VITE_BACKEND_URL=http://localhost:8000
```

Sideload the manifest into Excel via **Insert → Add-ins → Upload My Add-in**.

---

## Key Files

| File | Purpose |
|---|---|
| [src/services/agent.ts](src/services/agent.ts) | Agentic loop — prompt building, tool call parsing, iteration |
| [src/services/office.ts](src/services/office.ts) | All Office.js tool implementations |
| [src/services/chat.ts](src/services/chat.ts) | HTTP call to backend with auth token |
| [src/components/ChatComponent.tsx](src/components/ChatComponent.tsx) | Chat UI |
| [backend/index.js](backend/index.js) | Express server — JWT middleware + OpenAI proxy |
