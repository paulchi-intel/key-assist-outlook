# key-assist-outlook

An Outlook VBA macro that brings AI-powered email assistance directly into Microsoft Outlook.  
Supports both **ExpertGPT** and **GNAI** — Intel's internal AI gateways with OpenAI and Anthropic models.  
Select any email and trigger AI actions from the macro menu — no browser switching required.

---

## Features

### ⚡ Quick Actions (single selected email)
| Action | What it does |
|--------|-------------|
| **Summarize** | Condenses the email into bullet-point key concepts (Traditional Chinese) |
| **Translate** | Translates the email into clear, professional Traditional Chinese |
| **Action Items** | Extracts who needs to do what and by when |
| **Draft Reply** | Writes a friendly yet businesslike reply |
| **FAQ Convert** | Reformats the email as a Q&A FAQ |
| **Custom** | Enter any free-form instruction for the AI |

### 📚 FAQ Knowledge Base
- Reads all emails stored in an Outlook folder (`My Folders\FAQ`)
- Answers natural-language questions by searching across those emails
- Cites the specific source FAQ email(s) in its answer

### 📅 Daily Email Summary
- Summarize **today's** or **yesterday's** inbox in one click
- Or pick any **custom date** (absolute `YYYY-MM-DD` or relative `-7`)
- Output is organized by category, urgency, and includes a personal **Action Items** section

### ⚙️ Configuration
- GUI settings form (built with PowerShell/WinForms, no extra dependencies)
- Supports two Intel AI gateways:
  - **ExpertGPT** — enter a `pak_` API key; model list shows quota usage `(used/limit)`
  - **GNAI** — enter any non-`pak_` key; model list shows clean names without quota
- Models load automatically when you tab out of the API Key field, or click **Load Models**
- If the model API is unreachable, a built-in fallback list is shown
- Selected model and key are stored securely in the Windows Registry via `SaveSetting`

---

## Requirements

- Microsoft Outlook (desktop, Windows)
- An API key from either of the following Intel AI gateways:
  - **ExpertGPT** (`expertgpt.intel.com`) — keys start with `pak_`
  - **GNAI** (`gnai.intel.com`) — any non-`pak_` key
- Macros must be enabled in Outlook

---

## Installation

1. Open Outlook → press **Alt + F11** to open the VBA editor.
2. In the Project tree, select `ThisOutlookSession` or insert a new **Module**.
3. Paste the contents of `key-assist-outlook.vba` into the module.
4. Close the VBA editor and **enable macros** when prompted.
5. Run `ExpertGPT_Configure` (or any action macro) to set your API key and model.

> **Tip:** Add the macros to the Outlook Quick Access Toolbar or a custom Ribbon tab for one-click access.

---

## Configuration

| Setting | Details |
|---------|----------|
| **API Key** | ExpertGPT key (starts with `pak_`) **or** GNAI key (any other value). Stored in Windows Registry. |
| **Model** | Fetched live from the matching gateway. Falls back to a built-in list if the API is unreachable. Supports OpenAI and Anthropic (Claude) models. |
| **Quota display** | Shown as `model-name (used/limit)` for ExpertGPT keys only. GNAI shows plain model names. |
| **FAQ Folder** | Default path: `My Folders\FAQ`. Change `FAQ_FOLDER_PATH` constant in the source. |
| **Timeout** | OpenAI: 20 s · Anthropic: 120 s |

---

## Available Macros

```
ExpertGPT_Configure               — Open settings form
ExpertGPT_RefreshModelSelection   — Re-select model

ExpertGPT_AI_Summarize            — Summarize selected email
ExpertGPT_AI_Translate            — Translate selected email
ExpertGPT_AI_ActionItems          — Extract action items
ExpertGPT_AI_Reply                — Draft a reply
ExpertGPT_AI_FAQ                  — Convert to FAQ format
ExpertGPT_AI_Custom               — Custom prompt

ExpertGPT_FAQ_Ask                 — Ask a question against FAQ folder
ExpertGPT_SummarizeTodayEmails    — Summarize today's inbox
ExpertGPT_SummarizeYesterdayEmails — Summarize yesterday's inbox
ExpertGPT_SummarizeCustomDateEmails — Summarize inbox for a chosen date
```

---

## Notes

- AI results are opened as **new draft emails** addressed to yourself — nothing is sent automatically.
- The VBA code is plain text (`.vba`); import it into any standard Outlook module.
- The configuration form is rendered via an embedded PowerShell / WinForms script — no external dependencies or COM add-ins needed.
