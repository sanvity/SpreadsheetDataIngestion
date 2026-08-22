# Spreadsheet Data Extraction & Cell-Level Provenance Engine

Most spreadsheet tools assume you're working with clean, well-formed data. Real spreadsheets are rarely that polite. GHG accounting sheets, utility bills, and financial statements tend to arrive with scattered tables, inconsistent headers, and no predictable layout at all.

This engine is built for exactly that mess. Point it at an unstructured Excel workbook, ask a question in plain English, and it returns not just an answer, but the exact cell that answer came from — sheet, row, and column, down to the individual cell reference. No black box, no "trust me." Every number is traceable back to its source.

---

## How it works

### Finding tables without being told where they are

Rather than requiring a predefined range, the parser treats every populated cell as a node in a spatial grid graph and traverses it (DFS/BFS across all 8 neighboring cells) to identify distinct tables wherever they happen to live on the sheet.

Once the rough blocks are found, it merges the ones that clearly belong together — comparing bounding box overlap and checking how much empty space separates them — so a table that's been visually split doesn't get mistaken for two.

From there, it works out which rows are headers versus data by tracking where content shifts from text to numbers, handles headers spanning multiple rows, deduplicates column names, and assembles everything into a pandas DataFrame.

Underneath it all, the engine maintains a precise two-way map between DataFrame positions (`row_index`, `column_name`) and their original Excel coordinates (`sheet_name`, `excel_row`, `excel_col`, `excel_col_letter`). This mapping is what makes cell-level provenance possible later in the pipeline.

### A four-tier fallback for answering questions

Every query moves through a resolution cascade, escalating only when a step can't produce a confident answer:

1. **Intent parsing.** A language model — Mistral's `mistral-small-latest`, or a local Gemma 4 model via Ollama for a fully offline setup — converts the question into a structured spec: query type, metric, category, filters, aggregation.
2. **Sandboxed code generation.** The model writes pandas code to answer the query, which then runs inside a restricted AST sandbox. Only a safe whitelist of operations is permitted (`loc`, `iloc`, `groupby`, `agg`, `isin`); anything dangerous (`eval`, `exec`, `open`, `__import__`) is blocked outright.
3. **Schema-aware matching.** If that doesn't resolve cleanly, a heuristic resolver matches the query's entities and metrics directly against the DataFrame's rows and columns.
4. **Direct spatial scanning.** For sheets that aren't tabular at all — hierarchical GHG Scope 1/2/3 blocks, for instance — a resolver scans cells directly, locating metric values relative to nearby section headers.

### Provenance you can actually audit

Every answer is returned with its source, not just its value:

```text
0.253933
Source: Sheet[Stationary Energy], Row [8], Col [F] (Cell F8)
Metric Anchor: D8 | Category Anchor: C6
```

That traceability is the core design principle here. When a number feeds into a compliance report or a financial audit, "the model said so" isn't good enough — you need to be able to point to the exact cell it came from.

### An interface built to match

A Streamlit web app sits on top, with a glassmorphism design. Run a query and the matching cell lights up directly in the data preview (yellow background, red border) so you can see the answer in context immediately. Every interaction — query text, answer, timestamp, and provenance — is logged to a local SQLite database (`history.db`). Uploaded workbooks are hashed with SHA-256 so the same file never gets parsed twice.

---

## Architecture

```
                                 ┌─────────────────────────────────────────────────┐
                                 │              Streamlit Web Frontend             │
                                 │              (app.py & history.db)              │
                                 └────────────────────────┬────────────────────────┘
                                                          │
                                                          ▼
                                 ┌─────────────────────────────────────────────────┐
                                 │                SpreadsheetParser                │
                                 │  • 8-Neighbor Grid Graph Clustering             │
                                 │  • Spatial Bounding-Box Merging                 │
                                 │  • Multi-Level Header Detection                 │
                                 │  • DataFrame ↔ Excel Cell Coordinate Map        │
                                 └────────────────────────┬────────────────────────┘
                                                          │
                                                          ▼
                                 ┌─────────────────────────────────────────────────┐
                                 │               QueryPipeline Engine              │
                                 └────────────────────────┬────────────────────────┘
                                                          │
                 ┌────────────────────────────────────────┼────────────────────────────────────────┐
                 ▼                                        ▼                                        ▼
    ┌─────────────────────────┐              ┌─────────────────────────┐              ┌─────────────────────────┐
    │    LLM Query Manager    │              │ TableLLM & SafeSandbox │              │   Heuristic Resolvers   │
    │ Mistral AI / Gemma 4    │              │ AST Code Validation     │              │ DataFrame & Workbook    │
    │ Intent Parsing (JSON)   │              │ Sandboxed `exec`        │              │ Spatial Scanning        │
    └─────────────────────────┘              └─────────────────────────┘              └─────────────────────────┘
                                                          │
                                                          ▼
                                 ┌─────────────────────────────────────────────────┐
                                 │                ProvenanceMapper                 │
                                 │ Maps answer payload to Sheet, Row, Col Letter   │
                                 └─────────────────────────────────────────────────┘
```

**Stack**
- Python 3.10+
- pandas, openpyxl for data handling
- Mistral AI SDK (`mistral-small-latest`) and/or Ollama for local Gemma 4 (`gemma4:e2b`)
- Streamlit for the interface
- SQLite3 for query history (`history.db`)

---

## Getting started

### Install

```bash
git clone https://github.com/sanvijain/SpreadsheetExtraction.git
cd SpreadsheetExtraction
pip install -r requirements.txt
```

### API key (optional)

To use Mistral for intent parsing, set your key:

```bash
export MISTRAL_API_KEY="your-mistral-api-key"
```

No key? No problem. The engine falls back automatically to the local Gemma 4 model or the rule-based heuristic resolvers — it works fully offline if needed.

### Run the web app

```bash
streamlit run app.py
```

Then open `http://localhost:8501`.

### Or use the CLI

```bash
python pipeline.py --workbook "Sample Bill Sheet.xlsx" --query "what is Percapita PNG consumption in Residential"
```

```text
Final answer:
0.253933
Source: Sheet[Stationary Energy], Row [8], Col [F]

Stored JSON: /Users/sanvijain/SpreadsheetExtraction/query_result.json
Query history: /Users/sanvijain/SpreadsheetExtraction/query_history.json
```

---

## Repository structure

```text
├── pipeline.py                # Core backend: SpreadsheetParser, AST SafeCodeExecutor, Resolvers, QueryPipeline
├── app.py                     # Streamlit frontend with glassmorphism UI, cell highlighting, and DB history
├── requirements.txt           # Python package dependencies
├── history.db                 # SQLite database storing query audit logs
├── dummy_scattered.xlsx       # Sample scattered multi-table workbook for testing
├── Sample Bill Sheet.xlsx     # Sample utility & energy consumption workbook
├── GHG Accounting.xlsm        # Sample greenhouse gas accounting workbook
└── uploaded_files/            # Directory cache for user-uploaded Excel workbooks
```
