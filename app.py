from __future__ import annotations

import hashlib
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st

from pipeline import ExtractedTable, QueryPipeline, SpreadsheetParser, create_dummy_scattered_workbook


APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / "uploaded_files"
HISTORY_DB_PATH = APP_DIR / "history.db"


def init_db(db_path: Path = HISTORY_DB_PATH) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS query_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                query TEXT NOT NULL,
                answer TEXT NOT NULL,
                sheet_name TEXT,
                excel_row INTEGER,
                excel_col TEXT,
                cell_address TEXT,
                provenance_json TEXT NOT NULL
            )
            """
        )
        connection.commit()


def save_uploaded_file(uploaded_file) -> Path:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.sha256(file_bytes).hexdigest()[:12]
    destination = UPLOAD_DIR / f"{file_hash}_{uploaded_file.name}"
    destination.write_bytes(file_bytes)
    return destination


def parse_excel_to_tables(file_path: str | Path) -> Dict[str, ExtractedTable]:
    parser = SpreadsheetParser()
    return parser.parse(file_path)


def execute_table_llm_query(
    file_path: str | Path,
    query: str,
) -> Dict[str, Any]:
    pipeline = QueryPipeline(
        query_manager_provider="none",
        use_gemma_query_manager=False,
    )
    result = pipeline.answer_query(workbook_path=file_path, query=query)
    provenance = result["provenance"]
    return {
        "answer": result["answer"],
        "formatted_answer": result["formatted_answer"],
        "provenance": provenance,
    }


def insert_history_row(
    file_name: str,
    file_path: str,
    query: str,
    answer: Any,
    provenance: Dict[str, Any],
    db_path: Path = HISTORY_DB_PATH,
) -> None:
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            INSERT INTO query_history (
                created_at,
                file_name,
                file_path,
                query,
                answer,
                sheet_name,
                excel_row,
                excel_col,
                cell_address,
                provenance_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                file_name,
                file_path,
                query,
                str(answer),
                provenance.get("sheet_name"),
                provenance.get("excel_row"),
                provenance.get("excel_col_letter"),
                provenance.get("cell_address"),
                json.dumps(provenance),
            ),
        )
        connection.commit()


def fetch_query_history(db_path: Path = HISTORY_DB_PATH, limit: int = 50) -> List[Dict[str, Any]]:
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        rows = connection.execute(
            """
            SELECT created_at, file_name, query, answer, sheet_name, excel_row, excel_col, cell_address
            FROM query_history
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [dict(row) for row in rows]


def ensure_session_defaults() -> None:
    st.session_state.setdefault("uploaded_file_name", None)
    st.session_state.setdefault("uploaded_file_path", None)
    st.session_state.setdefault("parsed_tables", {})
    st.session_state.setdefault("extracted_tables", {})
    st.session_state.setdefault("chat_messages", [])
    st.session_state.setdefault("active_file_digest", None)
    st.session_state.setdefault("last_provenance", None)


def file_digest(file_path: str | Path) -> str:
    return hashlib.sha256(Path(file_path).read_bytes()).hexdigest()


def load_uploaded_workbook(uploaded_file) -> None:
    stored_path = save_uploaded_file(uploaded_file)
    current_digest = file_digest(stored_path)
    if st.session_state.get("active_file_digest") == current_digest:
        return

    extracted_tables = parse_excel_to_tables(stored_path)
    st.session_state.uploaded_file_name = uploaded_file.name
    st.session_state.uploaded_file_path = str(stored_path)
    st.session_state.extracted_tables = extracted_tables
    st.session_state.parsed_tables = {
        table_name: extracted_table.dataframe
        for table_name, extracted_table in extracted_tables.items()
    }
    st.session_state.chat_messages = []
    st.session_state.active_file_digest = current_digest
    st.session_state.last_provenance = None


def seed_demo_file() -> None:
    demo_path = APP_DIR / "dummy_scattered.xlsx"
    if not demo_path.exists():
        create_dummy_scattered_workbook(demo_path)
    extracted_tables = parse_excel_to_tables(demo_path)
    st.session_state.uploaded_file_name = demo_path.name
    st.session_state.uploaded_file_path = str(demo_path)
    st.session_state.extracted_tables = extracted_tables
    st.session_state.parsed_tables = {
        table_name: extracted_table.dataframe
        for table_name, extracted_table in extracted_tables.items()
    }
    st.session_state.chat_messages = []
    st.session_state.active_file_digest = file_digest(demo_path)
    st.session_state.last_provenance = None


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown("## Data Controls")
        uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
        if uploaded_file is not None:
            load_uploaded_workbook(uploaded_file)

        if st.button("Load Demo Workbook", use_container_width=True):
            seed_demo_file()

        st.markdown("---")
        st.markdown("## Query History")
        history_rows = fetch_query_history()
        if not history_rows:
            st.caption("No saved queries yet.")
        else:
            with st.expander("Recent Queries", expanded=True):
                for row in history_rows:
                    st.markdown(
                        "\n".join(
                            [
                                f"**{row['query']}**",
                                f"`{row['answer']}`",
                                f"{row['file_name']} • {row['sheet_name'] or 'Unknown sheet'}:{row['cell_address'] or '?'}",
                                f"<span style='color:#6b7280;font-size:0.8rem'>{row['created_at']}</span>",
                            ]
                        ),
                        unsafe_allow_html=True,
                    )
                    st.markdown("---")


def render_data_preview() -> None:
    st.markdown(
        """
        <div class="section-shell">
            <div class="section-title">Data Preview</div>
            <div class="section-subtitle">Detected table regions from the uploaded workbook.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.parsed_tables:
        st.info("Upload an `.xlsx` file from the sidebar to preview extracted tables.")
        return

    for table_name, extracted_table in st.session_state.extracted_tables.items():
        with st.container(border=True):
            st.markdown(f"**{table_name}**")
            preview_object = build_preview_object(extracted_table, st.session_state.last_provenance)
            st.dataframe(
                preview_object,
                use_container_width=True,
                height=min(360, 80 + len(extracted_table.dataframe) * 35),
            )


def build_preview_object(extracted_table: ExtractedTable, provenance: Optional[Dict[str, Any]]):
    dataframe = extracted_table.dataframe
    highlight_target = find_highlight_target(extracted_table, provenance)
    if highlight_target is None:
        return dataframe

    row_index, column_name = highlight_target

    def highlight_cell(value_frame: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=value_frame.index, columns=value_frame.columns)
        if row_index in styles.index and column_name in styles.columns:
            styles.loc[row_index, column_name] = (
                "background-color: rgba(255, 226, 89, 0.95); "
                "color: #111827; font-weight: 700; "
                "border: 3px solid #ef4444;"
            )
        return styles

    return dataframe.style.apply(highlight_cell, axis=None)


def find_highlight_target(
    extracted_table: ExtractedTable,
    provenance: Optional[Dict[str, Any]],
) -> Optional[tuple[int, str]]:
    if not provenance:
        return None

    table_name = provenance.get("table_name")
    if table_name and table_name != extracted_table.table_name:
        return None

    pandas_row_index = provenance.get("pandas_row_index")
    column_name = provenance.get("column_name")
    if pandas_row_index is not None and column_name in extracted_table.dataframe.columns:
        return int(pandas_row_index), str(column_name)

    excel_row = provenance.get("excel_row")
    excel_col_letter = provenance.get("excel_col_letter")
    if excel_row is None or excel_col_letter is None:
        return None

    for (row_index, candidate_column_name), coordinate in extracted_table.coordinate_map.items():
        if coordinate.excel_row == excel_row and coordinate.excel_col_letter == excel_col_letter:
            return int(row_index), str(candidate_column_name)

    return None


def render_chat() -> None:
    st.markdown(
        """
        <div class="section-shell chat-shell">
            <div class="section-title">Ask the Data</div>
            <div class="section-subtitle">Natural language queries with exact spreadsheet provenance.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    chat_container = st.container()
    with chat_container:
        recent_messages = st.session_state.chat_messages[-1:]
        for message in recent_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                provenance = message.get("provenance")
                if provenance:
                    render_provenance_box(provenance)

    query = st.chat_input("Ask a question about the uploaded workbook...")
    if not query:
        return

    if not st.session_state.uploaded_file_path:
        st.warning("Upload a workbook first so I have data to query.")
        return

    st.session_state.chat_messages.append({"role": "user", "content": query})
    with st.chat_message("user"):
        st.markdown(query)

    assistant_message: Dict[str, Any]
    with st.chat_message("assistant"):
        try:
            with st.spinner("Querying TableLLM..."):
                result = execute_table_llm_query(st.session_state.uploaded_file_path, query)
            answer_text = f"**Answer:** {result['answer']}"
            st.markdown(answer_text)
            render_provenance_box(result["provenance"])
            st.session_state.last_provenance = result["provenance"]
            assistant_message = {
                "role": "assistant",
                "content": answer_text,
                "provenance": result["provenance"],
            }
            insert_history_row(
                file_name=st.session_state.uploaded_file_name or "unknown.xlsx",
                file_path=st.session_state.uploaded_file_path,
                query=query,
                answer=result["answer"],
                provenance=result["provenance"],
            )
        except Exception as exc:
            error_text = f"**Query failed:** {exc}"
            st.error(error_text)
            assistant_message = {
                "role": "assistant",
                "content": error_text,
            }

    st.session_state.chat_messages.append(assistant_message)


def render_provenance_box(provenance: Dict[str, Any]) -> None:
    source_line = (
        f"Source: {provenance.get('sheet_name', 'Unknown sheet')}, "
        f"Row {provenance.get('excel_row', '?')}, "
        f"Col {provenance.get('excel_col_letter', provenance.get('excel_col', '?'))}"
    )
    cell_address = provenance.get("cell_address")
    metric_cell = provenance.get("metric_cell")
    category_cell = provenance.get("category_cell")

    details = [source_line]
    if cell_address:
        details.append(f"Cell: {cell_address}")
    if metric_cell:
        details.append(f"Metric anchor: {metric_cell}")
    if category_cell:
        details.append(f"Category anchor: {category_cell}")

    st.info("Provenance\n\n" + "\n\n".join(details))


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(201, 214, 255, 0.35), transparent 28%),
                radial-gradient(circle at top right, rgba(215, 244, 234, 0.45), transparent 30%),
                linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
        }
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 1400px;
        }
        .section-shell {
            margin-bottom: 0.8rem;
        }
        .section-title {
            font-size: 1.5rem;
            font-weight: 700;
            color: #0f172a;
            letter-spacing: -0.02em;
        }
        .section-subtitle {
            color: #475569;
            margin-top: 0.2rem;
            margin-bottom: 0.75rem;
        }
        .chat-shell {
            margin-top: 1.25rem;
            margin-bottom: 7rem;
        }
        [data-testid="stSidebar"] {
            background: rgba(255,255,255,0.7);
            backdrop-filter: blur(14px);
            border-right: 1px solid rgba(148, 163, 184, 0.2);
        }
        [data-testid="stSidebar"] * {
            color: #0f172a;
        }
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] div,
        [data-testid="stSidebar"] label {
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] code {
            color: #000000 !important;
            background: rgba(15, 23, 42, 0.08) !important;
        }
        div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stChatMessage"]) {
            padding-bottom: 0.5rem;
        }
        [data-testid="stChatMessage"] {
            background: rgba(255,255,255,0.74);
            border: 1px solid rgba(203, 213, 225, 0.8);
            border-radius: 18px;
            padding: 0.4rem 0.6rem;
            box-shadow: 0 12px 28px rgba(15, 23, 42, 0.04);
            margin-bottom: 0.75rem;
        }
        [data-testid="stChatMessage"] * {
            color: #0f172a !important;
            opacity: 1 !important;
        }
        [data-testid="stChatMessage"] p,
        [data-testid="stChatMessage"] li,
        [data-testid="stChatMessage"] span,
        [data-testid="stChatMessage"] div {
            color: #0f172a !important;
        }
        [data-testid="stChatInput"] textarea,
        [data-testid="stChatInput"] input {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            caret-color: #ffffff !important;
            font-weight: 500 !important;
        }
        [data-testid="stChatInput"] textarea::placeholder,
        [data-testid="stChatInput"] input::placeholder {
            color: rgba(255,255,255,0.72) !important;
            -webkit-text-fill-color: rgba(255,255,255,0.72) !important;
            opacity: 1 !important;
        }
    
        [data-testid="stChatInput"] {
            position: fixed;
            left: calc(50% + 2rem);
            transform: translateX(+40%);
            bottom: 1.5rem;
            width: min(38vw, 500px);
            max-width: 1400px;  /* match .block-container */
            z-index: 1000;
            background: rgba(15, 23, 42, 0.0);
        }
        
        [data-testid="stChatInput"] > div {
            background: rgba(15, 23, 42, 0.92) !important;
            border-radius: 18px !important;
            box-shadow: 0 18px 40px rgba(15, 23, 42, 0.22) !important;
        }
        [data-testid="stChatInput"] button {
            background: #ff4b4b !important;
            color: #ffffff !important;
        }
        @media (max-width: 1100px) {
            [data-testid="stChatInput"] {
                left: 1rem;
                right: 1rem;
                width: auto;
            }
        }
        [data-testid="stAlert"] * {
            opacity: 1 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Spreadsheet Data Ingestion",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_styles()
    init_db()
    ensure_session_defaults()


    # Force heading color to black
    st.markdown(
        """
        <style>
        h1 {
            color: black !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )


    # Add logo (local file or URL)
    st.image("transparent_logo.png", width=300)  # replace with your logo path or URL

    st.title("Spreadsheet Data Ingestion")
    st.caption("Upload unstructured Excel workbooks, inspect extracted tables, and query them with strict cell-level traceability.")

    render_sidebar()

    preview_col, chat_col = st.columns([1.1, 0.9], gap="large")
    with chat_col:
        render_chat()
    with preview_col:
        render_data_preview()


if __name__ == "__main__":
    main()
