# FrasaTwist — Comprehensive Technical Documentation

This document covers the architecture, API specifications, data schema, system workflows, and implementation details for every component of the **FrasaTwist** ecosystem.

---

## 📋 Table of Contents

1. [Project Overview](#1-project-overview)
2. [System Architecture & Component Relationships](#2-system-architecture--component-relationships)
3. [Module 1: Backend (`/backend`)](#3-module-1-backend-backend)
4. [Module 2: FrasaTwist Web (`/FrasaTwist-Web`)](#4-module-2-frasatwist-web-frasatwist-web)
5. [Module 3: Popup Web (`/popup-web`)](#5-module-3-popup-web-popup-web)
6. [Module 4: MS Word Add-in TaskPane (`/Test/src/taskpane`)](#6-module-4-ms-word-add-in-taskpane-testsrctaskpane)
7. [Environment Variable Configuration (.env)](#7-environment-variable-configuration-env)
8. [Running the System Locally](#8-running-the-system-locally)

---

## 1. Project Overview

**FrasaTwist** is an AI-powered academic and scientific writing assistant platform that integrates directly with Microsoft Word via an Office Add-in (TaskPane), backed by a robust REST API server and a public landing page.

### Vision & Problems Solved

In scientific writing (theses, dissertations, journals), authors frequently face:

1. **Loss of Formatting & Citations**: Generic paraphrase tools break automatic citations (Mendeley / Zotero), math equations (Word Math/OMML), and text formatting such as *italic* and **bold**.
2. **KBBI/EYD Compliance**: Foreign terms that must be italicized per EYD guidelines and non-standard spellings that are routinely missed.
3. **Seamless Integration**: The need to work without ever leaving Microsoft Word.

### Core Feature Highlights

| Feature | Description |
|---|---|
| **Format-Preserving Paraphrase Engine** | Isolates citations (`[[CIT_n]]`), math (`[[MATH_n]]`), and HTML format tags, paraphrases via Gemini 2.5 Flash, and reconstructs OOXML without breaking document structure. |
| **Auto-Italic Foreign Loanwords** | Automatically wraps foreign/technical terms with `<i>` tags in paraphrase output (Pro only). |
| **Typo & Grammar Corrector** | Audits 7 error categories mapped directly to Word document ranges using Word Bookmark Anchors. |
| **Advanced Grammar Check** | Pro-only hybrid pipeline: Hunspell generates candidates → Gemini LLM categorizes with a `source` field. |
| **Foreign Word Scanner** | Scans the full document for non-italic foreign words grouped by heading. Uses Hunspell without LLM. |
| **Interactive Diff Viewer** | Sentence-by-sentence before/after comparison with inline or stacked diff highlighting. |
| **Tier Management & Midtrans Payment** | Free (200 words/req, 20 req/day) vs Pro (unlimited, 750 words/req, full format preservation). |

---

## 2. System Architecture & Component Relationships

The FrasaTwist ecosystem consists of **4 main components**:

1. **`backend`** — Express.js REST API + PostgreSQL (Prisma ORM) + Google Gemini AI SDK.
2. **`FrasaTwist-Web`** — Public landing page & web portal (Next.js).
3. **`popup-web`** — Minimal Next.js app running inside the Office.js Dialog Popup for Google OAuth.
4. **`Test/src/taskpane`** — The MS Word Add-in (TaskPane UI) executed inside MS Word.

### System Architecture Diagram

```mermaid
graph TD
    subgraph "Client Layer"
        A["Microsoft Word"] -->|Runs Add-in| B["TaskPane (/Test/src/taskpane)"]
        C["Web Browser"] -->|Visits| D["FrasaTwist Web (/FrasaTwist-Web)"]
        B -->|Opens Dialog| E["Popup Web (/popup-web)"]
    end
    subgraph "Identity Provider"
        G["Google OAuth 2.0 API"]
    end
    subgraph "Payment Provider"
        H["Midtrans Payment Gateway"]
    end
    subgraph "Backend API (/backend)"
        F1["Auth & JWT Engine"]
        F2["Tier Enforcement & Quota"]
        F3["Format Preserving Engine"]
        F4["Typo Corrector Service"]
        F5["Billing & Webhook Handler"]
        F6["Foreign Word Scan (Hunspell)"]
        F7["Advanced Grammar Check"]
    end
    subgraph "AI & Database"
        I["Google Gemini AI (2.5 Flash / Flash Lite)"]
        J[("PostgreSQL (Supabase/Prisma)")]
    end
    E --> G --> E
    E -->|POST /auth/login| F1 --> J
    F1 --> E --> B
    B -->|POST /api/v2/paraphrase| F3
    B -->|POST /api/v2/correct-spelling| F4
    B -->|POST /api/v2/foreign-word-scan| F6
    B -->|POST /api/v2/advanced-grammar-check| F7
    F3 & F4 --> F2 --> J
    F3 & F4 & F7 --> I
    D -->|POST /api/v2/billing/create-transaction| F5 --> H
    H -->|Webhook| F5 --> J
```

### Cross-Module Integration Matrix

| Source | Target | Method | Purpose |
|:---|:---|:---|:---|
| TaskPane | popup-web | `Office.context.ui.displayDialogAsync` | Open isolated Google auth popup |
| popup-web | backend | `POST /auth/login` | Verify Google token, get JWT |
| popup-web | backend | `POST /auth/pending/:sessionId` | Store auth for TaskPane polling |
| TaskPane | backend | `GET /auth/pending/:sessionId` | Poll for auth result |
| popup-web | TaskPane | `Office.context.ui.messageParent` | Send tokens back to Word |
| TaskPane | backend | `POST /api/v2/paraphrase` | AI paraphrase with format preservation |
| TaskPane | backend | `POST /api/v2/correct-spelling` | 7-category spelling audit |
| TaskPane | backend | `POST /api/v2/foreign-word-scan` | Hunspell foreign word detection |
| TaskPane | backend | `POST /api/v2/advanced-grammar-check` | Hunspell + LLM grammar check (Pro) |
| FrasaTwist-Web | backend | `POST /api/v2/billing/create-transaction` | Create Midtrans subscription |
| Midtrans | backend | `POST /api/v2/webhooks/midtrans` | Payment notification → upgrade tier |

### Core Workflows

#### 1. Authentication Flow

1. User clicks **"Login Google"** in TaskPane.
2. `auth.js` calls `displayDialogAsync` → opens `popup-web/?session_id=...`
3. `popup-web` loads GIS, user selects Google account.
4. Google returns `idToken` → `popup-web` sends to `POST /auth/login`.
5. Backend verifies token, upserts user in PostgreSQL, generates **Access Token JWT (15 min)** + **Refresh Token (30 days)** (stored as SHA-256 hash in DB).
6. `popup-web` calls `messageParent` and `POST /auth/pending/:sessionId`.
7. TaskPane catches `DialogMessageReceived`, stores tokens in memory + `roamingSettings`.

#### 2. Format-Preserving Paraphrase Flow

1. User selects text → clicks **"Paraphrase"**.
2. `ooxmlParser.js` runs 6 parsing passes (see Module 4).
3. `taggedText` sent to `POST /api/v2/paraphrase` with `X-Device-Id` / `Authorization`.
4. `enforceTierRules` middleware checks quota; Gemini 2.5 Flash generates paraphrase.
5. Free Tier: HTML stripped → plain text. Pro Tier: full HTML preserved.
6. `diffViewer.js` shows before/after. On approval, `buildOoxmlFromTaggedText` reconstructs and inserts OOXML.

#### 3. Foreign Word Scan Flow

1. User opens **"Pindai Istilah Asing"** panel.
2. `foreignWordScanner.js` reads paragraphs, filters excluded sections, groups by heading.
3. `fullText` sent to `POST /api/v2/foreign-word-scan` → Hunspell detection (no LLM).
4. TaskPane checks italic status per candidate via `body.search()` + `font.italic`.
5. Results split into `notItalic` / `alreadyItalic`, rendered in `foreignWordPanel.js`.

#### 4. Advanced Grammar Check Flow (Pro Only)

1. User selects text → triggers **"Advanced Grammar Check"**.
2. Backend runs `extractCandidates(taggedText, 'detailed')` via Hunspell.
3. Candidate hint list (max 3 suggestions/word) sent to Gemini 2.5 Flash Lite with original text.
4. LLM categorizes across 7 error types, assigns `source: "hunspell_candidate"` or `"llm_detected"`.

---

## 3. Module 1: Backend (`/backend`)

### Tech Stack & Dependencies

| Dependency | Version | Purpose |
|---|---|---|
| Node.js ES Modules | — | Runtime |
| Express.js | `^5.2.1` | HTTP Framework |
| Prisma ORM | `^7.4.2` | PostgreSQL ORM |
| `@google/genai` | `^1.41.0` | Gemini AI SDK (primary) |
| `@google/generative-ai` | `^0.24.1` | Gemini AI SDK (legacy compat.) |
| `midtrans-client` | `^1.4.3` | Payment Gateway |
| `jsonwebtoken` | `^9.0.3` | JWT Auth |
| `google-auth-library` | `^10.9.1` | Google OAuth2 Token Verification |
| `pino` | `^10.3.1` | Structured JSON Logger |
| `express-rate-limit` | `^8.6.1` | IP Rate Limiter |
| `nspell` + `dictionary-en` | — | Hunspell spell-check (en_US) |
| `pg` | `^8.19.0` | PostgreSQL driver |

### Directory Structure (Backend)

```
backend/
├── prisma/
│   ├── migrations/                # PostgreSQL migration history
│   └── schema.prisma              # Prisma schema (User, Quota, Payment, Log)
├── services/
│   ├── dictionaryService.js       # Loads Hunspell id_ID + en_US dictionaries
│   └── hunspellCandidateExtractor.js  # Extracts foreign/misspelled word candidates
├── utils/
│   ├── advancedGrammarCheck.js    # Advanced Grammar Check: Hunspell + Gemini hybrid
│   ├── auth.js                    # JWT generation/validation, Tier Enforcement, Device Quota
│   ├── logger.js                  # Structured JSON Logger (Pino)
│   ├── midtrans.js                # Midtrans Snap init & SHA-512 Signature Validator
│   ├── prisma.js                  # Singleton Prisma client instance
│   ├── prompt.js                  # System Prompt Builder for Gemini AI Paraphrase
│   ├── typoCorrector.js           # Typo & Foreign Word Corrector (Gemini 2.5 Flash Lite)
│   └── util.js                    # Paraphrase service wrapper & helpers
├── .env                           # Environment variables (do NOT commit)
├── .env.example                   # Safe env template
├── Dockerfile                     # Docker / Cloud Run config
├── server.js                      # Entry point: all routes, middleware
├── package.json
└── vercel.json                    # Optional Vercel serverless config
```

### Database Schema (Prisma ORM)

File: [`backend/prisma/schema.prisma`](file:///c:/Kerja/FrasaTwist/backend/prisma/schema.prisma)

```prisma
enum Tier { FREE  PRO }
enum PaymentEventStatus { PENDING  SUCCESS  FAILED }

model User {
  id               String         @id @default(uuid())
  email            String         @unique
  fullName         String?
  tier             Tier           @default(FREE)
  subscriptionEnd  DateTime?
  createdAt        DateTime       @default(now())
  updatedAt        DateTime       @updatedAt
  refreshTokens    RefreshToken[]
  paymentEvents    PaymentEvent[]
  @@map("users")
}

model RefreshToken {
  id         String   @id @default(uuid())
  userId     String
  user       User     @relation(fields: [userId], references: [id], onDelete: Cascade)
  tokenHash  String   @unique   // SHA-256 hash of the plaintext token
  revoked    Boolean  @default(false)
  createdAt  DateTime @default(now())
  expiresAt  DateTime
  @@map("refresh_tokens")
}

// deviceId is namespaced:
//   Paraphrase/Spell:  raw deviceId (X-Device-Id header)
//   Foreign Word Scan: "fws:{deviceId}"
//   Advanced Grammar:  "agc:{userId}"
model QuotaUsage {
  id         String @id @default(uuid())
  deviceId   String
  usageDate  String // YYYY-MM-DD (WIB/Asia Jakarta)
  count      Int    @default(0)
  @@unique([deviceId, usageDate])
  @@map("quota_usage")
}

model PaymentEvent {
  id              String             @id @default(uuid())
  orderId         String             @unique
  userId          String
  user            User               @relation(fields: [userId], references: [id], onDelete: Cascade)
  grossAmount     Int
  status          PaymentEventStatus @default(PENDING)
  midtransStatus  String?
  paymentType     String?
  externalEventId String?            @unique
  rawPayload      Json?
  createdAt       DateTime           @default(now())
  updatedAt       DateTime           @updatedAt
  @@index([userId])
  @@map("payment_events")
}

model SecurityEvent {
  id          String   @id @default(uuid())
  eventType   String   // e.g. "MIDTRANS_INVALID_SIGNATURE", "MALFORMED_ORDER_ID"
  description String   @db.Text
  ip          String?
  payload     String?  @db.Text
  detail      Json?
  createdAt   DateTime @default(now())
  @@map("security_events")
}

model ParaphraseLog {
  id           String   @id @default(uuid())
  errorLog     String?  @db.Text
  iterations   Int      @default(1)
  isSuccessful Boolean  @default(true)
  mode         String?  // "formal" | "academic"
  createdAt    DateTime @default(now())
}
```

> **QuotaUsage Namespacing**: The `deviceId` column is reused across three quota features by prefixing the key:
> - Paraphrase / Spell Check: raw `{deviceId}` (from `X-Device-Id` header)
> - Foreign Word Scan: `fws:{deviceId}`
> - Advanced Grammar Check: `agc:{userId}`

---

### REST API Endpoints

#### Authentication

##### `POST /auth/login`
Verifies a Google OAuth ID Token and produces Access + Refresh Tokens.

- **Request Body**: `{ "idToken": "eyJhbGci..." }` (also accepts `credential` field)
- **Response 200**: `{ "accessToken", "refreshToken", "user": { id, email, fullName, tier } }`
- **Error 401**: Invalid or unverifiable token.

##### `POST /auth/refresh`
Rotates the Refresh Token, returns new Access + Refresh Tokens. Implements **Token Reuse Detection** — using a revoked token immediately revokes all sessions for that user.

- **Request Body**: `{ "refreshToken": "string" }`
- **Response 200**: `{ "accessToken", "refreshToken", "user": { ... } }`

##### `POST /auth/pending/:sessionId`
Called by `popup-web` after login to store auth data in an in-memory session store (TTL: 10 min). Enables TaskPane polling.

- **Request Body**: `{ "accessToken", "refreshToken", "user" }`
- **Response 200**: `{ "ok": true }`

##### `GET /auth/pending/:sessionId`
Polled by the TaskPane to retrieve the auth result. **One-time read** — entry deleted on retrieval.

- **Response** (ready): `{ "ready": true, "accessToken", "refreshToken", "user" }`
- **Response** (waiting): `{ "ready": false }`

##### `GET /auth/me`
Returns the authenticated user's profile from the database.

- **Headers**: `Authorization: Bearer <access_token>`
- **Response 200**: `{ "user": { id, email, fullName, tier, subscriptionEnd } }`

---

#### Core Feature Endpoints

##### `POST /api/v2/paraphrase`
Main endpoint for format-preserving AI paraphrase.

- **Headers**: `Authorization: Bearer <token>` (optional Free, required Pro) | `X-Device-Id` (required Free)
- **Request Body**:
  ```json
  {
    "originalText": "Plain text from Word",
    "taggedText": "Text with <b>HTML tags</b> and [[CIT_1]]",
    "objects": { "citations": [...], "math": [...] },
    "options": { "mode": "academic", "autoItalicForeign": true, "preserveFormatting": true }
  }
  ```
- **Response 200**:
  ```json
  { "success": true, "data": { "paraphrasedHtml": "...", "logId": "uuid" } }
  ```
- **Tier behavior**: Free Tier → HTML stripped to plain text. Pro Tier → full HTML preserved.

##### `POST /api/v2/correct-spelling`
7-category spelling and language audit.

- **Request Body**: `{ "taggedText": "string", "objects": { ... } }`
- **Response 200**:
  ```json
  {
    "corrections": [
      { "id": "corr_1", "type": "FOREIGN_WORD", "original": "framework",
        "replacementText": "framework", "applyItalic": true, "occurrence": 1,
        "explanation": "Istilah asing 'framework' perlu dicetak miring." }
    ],
    "count": 1
  }
  ```

**Correction Categories:**

| Category | Description |
|---|---|
| `TYPO` | Keystroke errors (e.g. `mengarahan` → `mengarahkan`) |
| `KATA_TIDAK_BAKU` | Non-standard word per KBBI/PUEBI (e.g. `aktifitas` → `aktivitas`) |
| `GRAMMAR` | Grammatical / phrase structure errors |
| `PUNGTUASI` | Punctuation errors per PUEBI |
| `KAPITALISASI` | Incorrect capitalization |
| `REDUNDANSI` | Pleonasm / redundant words |
| `FOREIGN_WORD` | Foreign terms that must be italicized per EYD |

##### `POST /api/v2/foreign-word-scan`
Scans document text for foreign word candidates using Hunspell (no LLM).

- **Headers**: `X-Device-Id` (required for Free Tier)
- **Request Body**: `{ "fullText": "string" }`
- **Tier Limits**: Free — max 2,500 words/scan, max 20 scans/day. Pro — no limits.
- **Response 200**:
  ```json
  {
    "candidates": [{ "word": "framework", "wordIndex": 42, "occurrenceIndex": 0 }],
    "count": 1, "wordCountChecked": 1543,
    "scansUsedToday": 3, "scansRemainingToday": 17
  }
  ```
  > For Pro users, quota fields are `null`.

##### `POST /api/v2/advanced-grammar-check`
**Pro-only.** Hybrid Hunspell + LLM grammar check.

- **Headers**: `Authorization: Bearer <pro_token>` (required)
- **Request Body**: `{ "taggedText": "string", "objects": { ... } }`
- **Pro Limits**: max 1,000 words/request, max 300 requests/day (quota key: `agc:{userId}`)
- **Response 200**:
  ```json
  {
    "corrections": [
      { "id": "corr_1", "type": "KATA_TIDAK_BAKU", "original": "merubah",
        "replacementText": "mengubah", "applyItalic": false, "occurrence": 1,
        "explanation": "Bentuk baku 'mengubah'.", "source": "hunspell_candidate" },
      { "id": "corr_2", "type": "GRAMMAR", "original": "di karenakan",
        "replacementText": "dikarenakan", "applyItalic": false, "occurrence": 1,
        "explanation": "Penulisan 'di' harus disambung.", "source": "llm_detected" }
    ],
    "count": 2, "wordCountChecked": 312,
    "requestsUsedToday": 1, "requestsRemainingToday": 299
  }
  ```
  > `source: "hunspell_candidate"` — flagged by Hunspell. `source: "llm_detected"` — found independently by LLM.

---

#### Billing & Webhooks

##### `POST /api/v2/billing/create-transaction`
Creates a Midtrans Snap transaction for Pro subscription.

- **Headers**: `Authorization: Bearer <access_token>`
- **Response 200**:
  ```json
  { "token": "snap-token", "redirectUrl": "https://...", "orderId": "FRASATWIST-PRO-{userId}-{ts}", "price": 49000 }
  ```

##### `POST /api/v2/webhooks/midtrans`
Webhook called by Midtrans on payment status change.

- **Security**: SHA-512 signature validation: `SHA512(order_id + status_code + gross_amount + ServerKey)`.
- **Idempotency**: Duplicate notifications for already-final events (`SUCCESS`/`FAILED`) are silently ignored.
- **Atomic Transaction**: Prisma `$transaction` simultaneously updates `PaymentEvent.status → SUCCESS` + `User.tier → PRO` + `subscriptionEnd + 30 days`.
- **Malformed order_id**: Logged to `SecurityEvent` and rejected with 400.

##### `POST /api/log-error`
Logs a frontend error message to `ParaphraseLog`.

- **Request Body**: `{ "errorMessage": "string" }`
- **Response 200**: `{ "success": true, "id": "log-uuid" }`

---

### Authentication, Rate Limiting & Tier Enforcement

#### IP Rate Limiter
Applied to all core feature endpoints. Limits unauthenticated/Free users to **10 req/min per IP**. Pro users with valid Bearer tokens are **exempted**.

#### Request Timeout
All requests time out after **60 seconds** → `408 Request Timeout`.

#### `enforceTierRules` Middleware

Applied to `/api/v2/paraphrase` and `/api/v2/correct-spelling`:

| Tier | Word Limit | Daily Cap | `X-Device-Id` Required |
|---|---|---|---|
| **FREE** | 200 words/req | 20 req/day (per Device ID) | ✅ Yes |
| **PRO** | 750 words/req | None | ❌ No |

Daily quota tracked in `QuotaUsage` table (date in WIB/Asia Jakarta timezone).

#### JWT Token Lifecycle

| Token | TTL | Client Storage | Notes |
|---|---|---|---|
| **Access Token** | 15 min | In-memory JS variable | HS256 signed with `JWT_ACCESS_SECRET` |
| **Refresh Token** | 30 days | `roamingSettings` + `localStorage` | Stored as SHA-256 hash in DB; rotated on each use |

**Token Reuse Detection**: Using a revoked refresh token triggers immediate revocation of all tokens for that user (force-logout all sessions).

---

### Payment Gateway Integration (Midtrans)

- **SDK**: `midtrans-client` Node.js SDK.
- **Order ID Format**: `FRASATWIST-PRO-{userId}-{timestamp}` — `userId` extracted via regex for atomic updates.
- **Sandbox/Production**: Controlled by `MIDTRANS_IS_PRODUCTION` env var.

---

## 4. Module 2: FrasaTwist Web (`/FrasaTwist-Web`)

### Tech Stack & Architecture (Web)

| Technology | Version | Purpose |
|---|---|---|
| Next.js | `v14.2.16` | App Router framework |
| TailwindCSS | `v3.4.1` | Styling |
| `lucide-react` | `v0.454.0` | Icons |
| Vercel | — | Hosting |

### Directory Structure & UI Components

```
FrasaTwist-Web/
├── app/
│   ├── globals.css           # Custom utilities & Tailwind color tokens
│   ├── layout.tsx            # Root layout (Metadata, Fonts: Inter/Outfit)
│   └── page.tsx              # Main landing page composition
├── components/
│   ├── Navbar.tsx            # Navigation bar with user profile & auth trigger
│   ├── Hero.tsx              # Hero section (headline, value prop, CTA)
│   ├── InteractiveDemo.tsx   # Live paraphrase simulation for visitors
│   ├── Features.tsx          # Feature showcase cards
│   ├── PricingSection.tsx    # FREE vs PRO comparison + Midtrans Snap integration
│   ├── AuthModal.tsx         # Sign in with Google modal (web context)
│   ├── InstallationGuide.tsx # Step-by-step MS Word Add-in installation guide
│   ├── FaqSection.tsx        # Accordion FAQ
│   └── Footer.tsx            # Footer links & copyright
└── lib/
    └── api.ts                # API client fetcher for backend communication
```

---

## 5. Module 3: Popup Web (`/popup-web`)

### Purpose & Role (Office Dialog Context)

The MS Word Add-in runs inside a sandboxed IFrame (CORS, COOP, sandbox restrictions). Google OAuth 2.0 cannot be initiated inside the TaskPane directly.

**`popup-web`** is a minimal, standalone Next.js app opened via `Office.context.ui.displayDialogAsync()`. Its **only purpose** is to handle Google Sign-In and relay the result back to the TaskPane.

```
MS Word TaskPane
  → displayDialogAsync(popup-web URL)
        ↓
Popup Dialog Window (popup-web)
  1. Load Google Identity Services
  2. User selects Google account → idToken
  3. POST /auth/login → backend
  4. POST /auth/pending/:sessionId → backend (for polling)
  5. messageParent(authData) → TaskPane
        ↓
TaskPane (DialogMessageReceived)
  1. Parse authData
  2. Store tokens (memory + roamingSettings)
  3. dialog.close()
  4. Update UI
```

> **Why no Client Secret?** GIS uses a client-side ID token that the backend verifies directly against Google's public keys — no server-to-server exchange needed. `popup-web` never stores a Client Secret.

### File Structure & Modules

```
popup-web/
├── app/
│   ├── layout.tsx        # Minimal layout (no Office.js here)
│   ├── page.tsx          # Single page: GIS login + backend handshake
│   └── globals.css       # Minimal styling (no Tailwind)
├── lib/
│   ├── officeMessage.ts  # Dynamic Office.js load + messageParent
│   └── googleIdentity.ts # Dynamic GIS load
├── next.config.mjs
├── package.json
└── .env.example
```

### MessageParent Communication Protocol

```typescript
// popup-web/lib/officeMessage.ts
export function sendResultToTaskPane(data: {
  accessToken: string;
  refreshToken: string;
  user: Record<string, unknown>;
}): boolean {
  if (!isOfficeDialogContext()) return false;
  const w = window as any;
  w.Office.context.ui.messageParent(JSON.stringify(data));
  return true;
}
```

---

## 6. Module 4: MS Word Add-in TaskPane (`/Test/src/taskpane`)

### Tech Stack & Architecture (Add-in)

| Technology | Purpose |
|---|---|
| MS Office Add-in JS SDK (`office-js`) | Office.js API access |
| Webpack 5 + Babel Loader | Bundler & transpiler |
| ES6 Modules | Modular client architecture |
| Vanilla CSS (Custom Properties) | Design tokens & styling |

### File Structure & Components

```
Test/src/taskpane/
├── auth.js               # Dialog launcher, token storage (roamingSettings/localStorage), refresh
├── config.js             # Server endpoints & WordprocessingML XML boilerplates
├── diffViewer.js         # HTML format-preserving diff renderer & sentence segmentation
├── foreignWordPanel.js   # UI rendering for Foreign Word Scan results (grouped by heading)
├── foreignWordScanner.js # Scan orchestrator: paragraph extraction → API → italic check
├── headingGrouper.js     # Pure utility: group paragraphs by nearest heading
├── ooxmlParser.js        # Pure OOXML Parser (Pass 1–6) & Reconstruction Engine
├── paraphrase.js         # Office.js selection wrapper & API integration
├── scanExclusions.js     # Filter: excludes bibliography/glossary/appendix headings
├── state.js              # Central reactive state store (user, selections, corrections)
├── taskpane.css          # TaskPane UI styling (light/dark mode)
├── taskpane.html         # TaskPane HTML markup
├── taskpane.js           # Main entry point (Office.onReady, UI event listeners)
├── tokens.css            # Color & spacing design tokens
├── typoCorrector.js      # Typo audit engine with Word Bookmark Region Anchoring
└── uiHelpers.js          # Toast notifications, modal management, loading indicators
```

---

### Format-Preserving Engine (OOXML Parser)

File: [`ooxmlParser.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/ooxmlParser.js) — a **pure engine** (no DOM dependencies) for extracting and reconstructing MS Word OOXML.

#### 6 Parsing Passes (`parseOoxmlToIntermediateHtml`)

| Pass | Target | Action |
|---|---|---|
| **Pass 1** | `<m:oMath>` | Replaces math XML with `[[MATH_n]]` placeholders; saves originals in `mappingData.math`. |
| **Pass 2** | `<w:fldSimple>`, `<w:instrText>` | Detects Zotero/Mendeley/Word citations → `[[CIT_n]]`; saves in `mappingData.citations`. |
| **Pass 3** | `<w:b>`, `<w:i>`, `w:pStyle` | Converts bold/italic/Heading runs → `<b>`, `<i>`, `<h1>`–`<h3>`. |
| **Pass 4** | `<w:pPr>` | Saves paragraph formatting (alignment, spacing, indentation) to `mappingData.paragraphFormatting`. |
| **Pass 5** | XML entities | Decodes `&lt;`, `&gt;`, `&amp;` into LLM-readable text. |
| **Pass 6** | Content Controls | Isolates additional Mendeley Content Control objects. |

#### Reconstruction Engine (`buildOoxmlFromTaggedText`)

Reverses the parsing process after LLM returns `paraphrasedHtml`:

1. Splits output by paragraph.
2. Restores `[[CIT_n]]` / `[[MATH_n]]` tokens → original XML blocks from `mappingData`.
3. Converts `<b>`, `<i>`, `<h1>`–`<h3>` → `<w:r><w:rPr>...</w:rPr><w:t>...</w:t></w:r>` OOXML.
4. Assembles full `pkg:package` XML → inserts via `selection.insertOoxml()`.

---

### Diff Viewer & Sentence Tracker

File: [`diffViewer.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/diffViewer.js)

- Segments both texts into sentences.
- Computes word-level diffs per sentence pair.
- Renders inline diff (strikethrough red / inserted green) or stacked diff (two-panel).
- Preserves HTML tags (`<b>`, `<i>`, etc.) throughout the diff process.

---

### Typo & Foreign Word Corrector Engine (TaskPane)

File: [`typoCorrector.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/typoCorrector.js)

#### Word Bookmark Region Anchoring (`frasatwist_typo_region`)

Prevents modifications outside the user's selection:

1. On **"Cek Ejaan"** click → selected range is bookmarked as `frasatwist_typo_region`.
2. When a correction is applied → `searchScope.search()` runs **only within the bookmarked range**.
3. For `FOREIGN_WORD` with `applyItalic: true` → `range.font.italic = true` applied without disturbing surrounding citations.

---

### Foreign Word Scanner Module

File: [`foreignWordScanner.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/foreignWordScanner.js)

#### Scan Orchestration Flow

```
runForeignWordScan()
  ├─ 1. Read document.body.paragraphs (Office.js)
  ├─ 2. filterExcludedParagraphs()  → remove bibliography/glossary/appendix
  ├─ 3. groupParagraphsByHeading()  → Map<headingLabel, paragraph[]>
  ├─ 4. Count words; enforce Free tier 2,500 word limit
  ├─ 5. Concat fullText → POST /api/v2/foreign-word-scan
  ├─ 6. For each candidate: body.search(word) → check font.italic
  ├─ 7. Split → notItalic[] vs alreadyItalic[]
  └─ 8. renderForeignWordScanResults()  [foreignWordPanel.js]
```

#### Supporting Modules

| Module | Purpose |
|---|---|
| [`headingGrouper.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/headingGrouper.js) | Pure utility (no Office.js). Groups paragraphs by nearest preceding heading using regex `/heading\|judul\|title\|bab/i`. Pre-heading paragraphs fall under `"Bagian Awal"`. |
| [`scanExclusions.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/scanExclusions.js) | Filters out paragraphs from excluded sections (bibliography, glossary, appendix) based on heading names. |
| [`foreignWordPanel.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/foreignWordPanel.js) | Renders scan results UI grouped by document heading, with per-word italic-application controls. |

---

### Auth Session Management & Roaming Settings

File: [`auth.js`](file:///c:/Kerja/FrasaTwist/Test/src/taskpane/auth.js)

| Storage | Data | Scope |
|---|---|---|
| **In-memory JS variable** | Access Token | Current session only; cleared on reload |
| **`Office.context.roamingSettings`** | Refresh Token | Persists across devices for same MS 365 account |
| **`localStorage`** | Refresh Token (fallback) | Browser-specific persistence |

**Auto-Refresh**: `fetchWithAuth()` wraps every API call. On `401 Unauthorized`, it automatically calls `POST /auth/refresh`, updates tokens in storage, and retries the original request once.

---

## 7. Environment Variable Configuration (.env)

### Backend (`/backend/.env`)

```env
PORT=8000
DATABASE_URL="postgresql://user:password@db-host:5432/frasatwist?schema=public"
DIRECT_URL="postgresql://user:password@db-host:5432/frasatwist?schema=public"

GEMINI_API_KEY="AIzaSy..."
JWT_ACCESS_SECRET="super-secret-jwt-key-frasatwist-2026"
GOOGLE_CLIENT_ID="531010905615-....apps.googleusercontent.com"

# Optional: Google Cloud Vertex AI routing
GOOGLE_CLOUD_PROJECT="your-project-id"
GOOGLE_CLOUD_LOCATION="global"

MIDTRANS_SERVER_KEY="SB-Mid-server-..."
MIDTRANS_CLIENT_KEY="SB-Mid-client-..."
MIDTRANS_IS_PRODUCTION=false
PRO_PRICE_IDR=49000
```

### FrasaTwist Web (`/FrasaTwist-Web/.env.local`)

```env
NEXT_PUBLIC_BACKEND_URL="http://localhost:8000"
NEXT_PUBLIC_MIDTRANS_CLIENT_KEY="SB-Mid-client-..."
```

### Popup Web (`/popup-web/.env`)

```env
NEXT_PUBLIC_BACKEND_API_URL="http://localhost:8000"
NEXT_PUBLIC_GOOGLE_CLIENT_ID="531010905615-....apps.googleusercontent.com"
```

---

## 8. Running the System Locally

### Step 1: Backend API Server

```bash
cd c:\Kerja\FrasaTwist\backend
npm install
npx prisma db push      # Sync schema to PostgreSQL
npm start               # http://localhost:8000
```

### Step 2: Popup Web Server

```bash
cd c:\Kerja\FrasaTwist\popup-web
npm install
cp .env.example .env.local   # Fill in backend URL & Google Client ID
npm run dev                  # http://localhost:3001
```

### Step 3: FrasaTwist Web Landing Page

```bash
cd c:\Kerja\FrasaTwist\FrasaTwist-Web
npm install
npm run dev             # http://localhost:3000
```

### Step 4: MS Word Add-in (TaskPane)

```bash
cd c:\Kerja\FrasaTwist\Test
npm install
npm run start           # Webpack dev server + registers manifest.xml in MS Word
```

### Google Cloud Console Setup (Required for OAuth)

For `popup-web` login to work, the OAuth 2.0 Client ID must have:

- **Authorized JavaScript Origins**:
  - `http://localhost:3001` (local dev)
  - `https://<your-popup-web-domain>` (production)
- **No Redirect URIs needed** — GIS uses a callback-based flow, not redirect.

---

*This documentation is maintained for architecture review, feature development, and onboarding new contributors to the FrasaTwist ecosystem.*
