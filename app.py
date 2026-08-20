from __future__ import annotations

from datetime import datetime
import hashlib
import hmac
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
POLLING_FILE = BASE_DIR / "Polling_Station_LGD mapping.xlsx"
LGD_FILE = BASE_DIR / "LGD with constituency.xlsx"
OUTPUT_PREFIX = "Polling_Station_LGD mapping_manual"

POLLING_SHEET = "Unmapped"
LGD_SHEET = "Sheet1"

POLLING_CONSTITUENCY_COL = "Constituency"
POLLING_VILLAGE_COL = "Village"
POLLING_DETAILS_COL = "Details of Sections covered by the part (Original)"

LGD_CONSTITUENCY_COL = "Assembly Constituency Name"
LGD_CONSTITUENCY_NUMBER_COL = "Assembly Constituency Name and Number"
LGD_VILLAGE_COL = "Village Name"
LGD_CODE_COL = "Village_Code"
LGD_BLOCK_COL = "Block Name"
LGD_DISTRICT_COL = "District Name"
LGD_ENTITY_COL = "Entity Type"

TARGET_COLUMNS = {
    "Matched LGD Village/Ward Name": LGD_VILLAGE_COL,
    "LGD Village Code": LGD_CODE_COL,
    "LGD Block Name": LGD_BLOCK_COL,
    "LGD District Name": LGD_DISTRICT_COL,
    "LGD Entity Type": LGD_ENTITY_COL,
}


st.set_page_config(
    page_title="Polling Station LGD Manual Mapper",
    page_icon=":material/table_view:",
    layout="wide",
)


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def configured_users() -> dict[str, str]:
    try:
        users = st.secrets["auth"]["users"]
    except KeyError:
        return {}
    return {username: details["password_hash"] for username, details in users.items()}


def require_login() -> None:
    users = configured_users()
    if not users:
        st.error("Login is not configured. Add users in Streamlit secrets.")
        st.code(
            "[auth.users.user1]\n"
            'password_hash = "sha256-password-hash"\n',
            language="toml",
        )
        st.stop()

    if st.session_state.get("authenticated"):
        with st.sidebar:
            st.write(f"Signed in as {st.session_state['username']}")
            if st.button("Sign out"):
                st.session_state.clear()
                st.rerun()
        return

    st.title("Polling Station LGD Manual Mapper")
    with st.form("login_form"):
        username = st.text_input("Login ID")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", type="primary")

    if submitted:
        expected_hash = users.get(username.strip())
        supplied_hash = hash_password(password)
        if expected_hash and hmac.compare_digest(expected_hash, supplied_hash):
            st.session_state.authenticated = True
            st.session_state.username = username.strip()
            st.rerun()
        else:
            st.error("Invalid login ID or password.")

    st.stop()


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def constituency_name(value: object) -> str:
    text = normalize_text(value)
    if "-" in text:
        return text.split("-", 1)[1].strip().upper()
    return text.upper()


@st.cache_data(show_spinner=False)
def load_workbooks(
    polling_source: str | bytes, lgd_source: str | bytes
) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    polling_sheets = pd.read_excel(polling_source, sheet_name=None)
    lgd = pd.read_excel(lgd_source, sheet_name=LGD_SHEET)
    return polling_sheets, lgd


def blank_mask(df: pd.DataFrame) -> pd.Series:
    cols = [col for col in TARGET_COLUMNS if col in df.columns]
    if not cols:
        return pd.Series(False, index=df.index)
    return df[cols].isna().all(axis=1) | df[cols].astype(str).apply(
        lambda row: all(item.strip() in {"", "nan", "None"} for item in row),
        axis=1,
    )


def build_excel(workbook: dict[str, pd.DataFrame], updated_unmapped: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, sheet_df in workbook.items():
            if sheet_name == POLLING_SHEET:
                updated_unmapped.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        mapped_now = updated_unmapped[~blank_mask(updated_unmapped)].copy()
        mapped_now.to_excel(writer, sheet_name="Manual_Mapped_Unmapped", index=False)

    return output.getvalue()


def format_lgd_option(row: pd.Series) -> str:
    code = normalize_text(row.get(LGD_CODE_COL))
    village = normalize_text(row.get(LGD_VILLAGE_COL))
    district = normalize_text(row.get(LGD_DISTRICT_COL))
    entity = normalize_text(row.get(LGD_ENTITY_COL))
    return f"{village} | LGD Code: {code} | {district} | {entity}"


def unique_clean_values(series: pd.Series) -> list[str]:
    return sorted(
        value
        for value in series.dropna().astype(str).map(str.strip).unique()
        if value
    )


def ensure_state(unmapped_df: pd.DataFrame) -> None:
    current_signature = st.session_state.get("source_signature")
    if current_signature != st.session_state.get("loaded_source_signature"):
        for key in ["working_unmapped", "mapping_log", "polling_table", "lgd_table"]:
            st.session_state.pop(key, None)
        st.session_state.loaded_source_signature = current_signature

    if "working_unmapped" not in st.session_state:
        st.session_state.working_unmapped = unmapped_df.copy()
    if "mapping_log" not in st.session_state:
        st.session_state.mapping_log = []


def apply_mapping(row_indices: list[int], lgd_row: pd.Series) -> None:
    for target_col, source_col in TARGET_COLUMNS.items():
        st.session_state.working_unmapped.loc[row_indices, target_col] = lgd_row.get(
            source_col
        )

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for row_index in row_indices:
        st.session_state.mapping_log.append(
            {
                "Mapped At": now,
                "Polling Row": int(row_index) + 2,
                "Constituency": st.session_state.working_unmapped.loc[
                    row_index, POLLING_CONSTITUENCY_COL
                ],
                "Polling Village": st.session_state.working_unmapped.loc[
                    row_index, POLLING_VILLAGE_COL
                ],
                "LGD Village/Ward Name": lgd_row.get(LGD_VILLAGE_COL),
                "LGD Village Code": lgd_row.get(LGD_CODE_COL),
            }
        )


require_login()

with st.sidebar:
    st.header("Workbook Source")
    polling_upload = st.file_uploader(
        "Polling mapping workbook",
        type=["xlsx"],
        help="Upload Polling_Station_LGD mapping.xlsx if it is not bundled on the server.",
    )
    lgd_upload = st.file_uploader(
        "LGD workbook",
        type=["xlsx"],
        help="Upload LGD with constituency.xlsx if it is not bundled on the server.",
    )

polling_source: str | bytes | None
lgd_source: str | bytes | None
polling_source = polling_upload.getvalue() if polling_upload else None
lgd_source = lgd_upload.getvalue() if lgd_upload else None

if polling_source is None and POLLING_FILE.exists():
    polling_source = str(POLLING_FILE)
if lgd_source is None and LGD_FILE.exists():
    lgd_source = str(LGD_FILE)

if polling_source is None or lgd_source is None:
    st.title("Polling Station LGD Manual Mapper")
    st.info("Upload both Excel workbooks from the sidebar to start mapping.")
    st.stop()

st.session_state.source_signature = (
    (polling_upload.name, polling_upload.size) if polling_upload else str(POLLING_FILE),
    (lgd_upload.name, lgd_upload.size) if lgd_upload else str(LGD_FILE),
)

try:
    polling_workbook, lgd_df = load_workbooks(polling_source, lgd_source)
except FileNotFoundError as exc:
    st.error(f"Missing workbook: {exc.filename}")
    st.stop()
except Exception as exc:
    st.error(f"Could not load the Excel files: {exc}")
    st.stop()

if POLLING_SHEET not in polling_workbook:
    st.error(f"Sheet '{POLLING_SHEET}' was not found in {POLLING_FILE.name}.")
    st.stop()

unmapped_source = polling_workbook[POLLING_SHEET]
ensure_state(unmapped_source)

working_unmapped = st.session_state.working_unmapped
unmapped_only = working_unmapped[blank_mask(working_unmapped)].copy()

st.title("Polling Station LGD Manual Mapper")

left_metric, mid_metric, right_metric = st.columns(3)
left_metric.metric("Total polling rows", f"{len(working_unmapped):,}")
mid_metric.metric("Still unmapped", f"{len(unmapped_only):,}")
right_metric.metric("Mapped in this session", f"{len(st.session_state.mapping_log):,}")

constituencies = sorted(
    value
    for value in unmapped_only[POLLING_CONSTITUENCY_COL].dropna().astype(str).unique()
    if value.strip()
)

if not constituencies:
    st.success("All rows in the Unmapped sheet have LGD mapping values.")
else:
    st.divider()

    filter_col_1, filter_col_2, filter_col_3 = st.columns([1, 1, 1])
    with filter_col_1:
        selected_constituency = st.selectbox(
            "Polling constituency",
            constituencies,
            index=0,
        )

    selected_name = constituency_name(selected_constituency)
    polling_view = unmapped_only[
        unmapped_only[POLLING_CONSTITUENCY_COL].astype(str) == selected_constituency
    ].copy()

    lgd_constituency_name = lgd_df[LGD_CONSTITUENCY_COL].map(constituency_name)
    lgd_constituency_number = lgd_df[LGD_CONSTITUENCY_NUMBER_COL].map(constituency_name)
    lgd_constituency_text = (
        lgd_df[LGD_CONSTITUENCY_COL].map(constituency_name)
        + " "
        + lgd_df[LGD_CONSTITUENCY_NUMBER_COL].map(constituency_name)
    )

    lgd_constituency_options = unique_clean_values(lgd_df[LGD_CONSTITUENCY_NUMBER_COL])
    default_lgd_index = 0
    for option_index, option in enumerate(lgd_constituency_options):
        if selected_name in constituency_name(option):
            default_lgd_index = option_index
            break

    with filter_col_2:
        selected_lgd_constituency = st.selectbox(
            "LGD constituency",
            lgd_constituency_options,
            index=default_lgd_index,
        )

    selected_lgd_name = constituency_name(selected_lgd_constituency)
    lgd_view = lgd_df[
        (lgd_constituency_name == selected_lgd_name)
        | (lgd_constituency_number == selected_lgd_name)
        | lgd_constituency_text.str.contains(selected_lgd_name, na=False)
    ].copy()

    district_options = ["All districts"] + unique_clean_values(lgd_view[LGD_DISTRICT_COL])
    polling_districts = unique_clean_values(polling_view["District"])
    default_district_index = 0
    for polling_district in polling_districts:
        if polling_district in district_options:
            default_district_index = district_options.index(polling_district)
            break

    with filter_col_3:
        selected_lgd_district = st.selectbox(
            "LGD district",
            district_options,
            index=default_district_index,
        )

    if selected_lgd_district != "All districts":
        lgd_view = lgd_view[
            lgd_view[LGD_DISTRICT_COL].astype(str) == selected_lgd_district
        ].copy()

    polling_col, lgd_col = st.columns([1.1, 1])

    with polling_col:
        st.subheader("Unmapped polling records")
        st.caption(f"{len(polling_view):,} unmapped rows in {selected_constituency}")
        polling_display = polling_view[
            [
                POLLING_CONSTITUENCY_COL,
                "District",
                "Electoral Roll Part No.",
                POLLING_VILLAGE_COL,
                POLLING_DETAILS_COL,
            ]
        ].copy()
        polling_display.insert(0, "Excel Row", polling_display.index + 2)
        selected_rows = st.dataframe(
            polling_display,
            hide_index=True,
            use_container_width=True,
            selection_mode="multi-row",
            on_select="rerun",
            key="polling_table",
        )

    with lgd_col:
        st.subheader("LGD village list")
        if lgd_view.empty:
            st.warning("No LGD records found for the selected filters.")
            search_source = lgd_view.copy()
        else:
            search_source = lgd_view.copy()

        search_text = st.text_input(
            "Search LGD village or code",
            placeholder="Type village name or LGD code",
        )
        if search_text.strip():
            needle = search_text.strip().lower()
            search_columns = [
                LGD_VILLAGE_COL,
                LGD_CODE_COL,
            ]
            search_blob = (
                search_source[search_columns]
                .astype(str)
                .agg(" ".join, axis=1)
                .str.lower()
            )
            search_source = search_source[search_blob.str.contains(needle, na=False)]

        lgd_options = list(search_source.index)
        if lgd_options:
            st.caption(f"{len(search_source):,} LGD villages/wards")
            lgd_display = search_source[
                [
                    LGD_VILLAGE_COL,
                    LGD_CODE_COL,
                    LGD_DISTRICT_COL,
                    LGD_ENTITY_COL,
                    LGD_CONSTITUENCY_NUMBER_COL,
                ]
            ].copy()
            lgd_display.columns = [
                "Village / Ward Name",
                "LGD Code",
                "District",
                "Type",
                "LGD Constituency",
            ]
            selected_lgd_rows = st.dataframe(
                lgd_display,
                hide_index=True,
                use_container_width=True,
                selection_mode="single-row",
                on_select="rerun",
                key="lgd_table",
            )
            picked_lgd_positions = selected_lgd_rows.selection.rows
            selected_lgd_index = (
                search_source.iloc[picked_lgd_positions[0]].name
                if picked_lgd_positions
                else None
            )
        else:
            selected_lgd_index = None
            st.info("No LGD records match the current search.")

    picked_positions = selected_rows.selection.rows if selected_rows else []
    picked_indices = polling_view.iloc[picked_positions].index.tolist()

    action_col, save_col = st.columns([1, 1])
    with action_col:
        st.caption(
            f"Selected polling rows: {len(picked_indices)} | "
            f"Selected LGD: {format_lgd_option(search_source.loc[selected_lgd_index]) if selected_lgd_index is not None else 'None'}"
        )
        if st.button(
            "Apply selected LGD mapping",
            type="primary",
            disabled=not picked_indices or selected_lgd_index is None,
            use_container_width=True,
        ):
            apply_mapping(picked_indices, search_source.loc[selected_lgd_index])
            st.success(f"Mapped {len(picked_indices)} polling row(s).")
            st.rerun()

    with save_col:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"{OUTPUT_PREFIX}_{timestamp}.xlsx"
        excel_bytes = build_excel(polling_workbook, st.session_state.working_unmapped)
        st.download_button(
            "Download updated Excel",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if st.button("Save updated Excel in this folder", use_container_width=True):
            output_dir = BASE_DIR / "outputs"
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / output_name
            output_path.write_bytes(excel_bytes)
            st.success(f"Saved outputs/{output_path.name}")

if st.session_state.mapping_log:
    st.divider()
    st.subheader("Session mapping log")
    st.dataframe(pd.DataFrame(st.session_state.mapping_log), hide_index=True)

with st.sidebar:
    st.header("Files")
    st.write(polling_upload.name if polling_upload else POLLING_FILE.name)
    st.write(lgd_upload.name if lgd_upload else LGD_FILE.name)
    if st.button("Reload original workbooks"):
        load_workbooks.clear()
        for key in ["working_unmapped", "mapping_log", "polling_table", "lgd_table"]:
            st.session_state.pop(key, None)
        st.rerun()
