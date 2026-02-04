import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# Try to import RapidFuzz
try:
    from rapidfuzz import process as rf_process, fuzz as rf_fuzz
    HAS_RAPIDFUZZ = True
except Exception:
    HAS_RAPIDFUZZ = False
    import difflib

# ---------------------------
# Utilities
# ---------------------------
def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    s = re.sub(r'[\s\-\_]+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s

def fuzzy_best_match_rapidfuzz(q, choices):
    # returns (best_choice, score)
    best = rf_process.extractOne(q, choices, scorer=rf_fuzz.token_sort_ratio)
    if best:
        return best[0], int(best[1])
    return None, 0

def fuzzy_best_match_difflib(q, choices):
    matches = difflib.get_close_matches(q, choices, n=1, cutoff=0.0)
    if matches:
        score = int(difflib.SequenceMatcher(None, q, matches[0]).ratio() * 100)
        return matches[0], score
    return None, 0

def best_match(query, choices, method='rapidfuzz'):
    q = normalize_text(query)
    if not q:
        return None, 0
    if method == 'exact':
        # exact normalized match
        if q in choices:
            return q, 100
        return None, 0
    if method == 'rapidfuzz' and HAS_RAPIDFUZZ:
        return fuzzy_best_match_rapidfuzz(q, choices)
    # fallback to difflib
    return fuzzy_best_match_difflib(q, choices)

def load_file(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def create_excel_bytes(df_missing, df_matches):
    # Create an Excel file in memory with two sheets
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_matches.to_excel(writer, sheet_name='Matches', index=False)
        df_missing.to_excel(writer, sheet_name='Missing', index=False)
        workbook = writer.book
        # simple formatting
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white'})
        for sheet in ['Matches', 'Missing']:
            worksheet = writer.sheets[sheet]
            worksheet.set_column(0, len(df_matches.columns), 25)
            for col_num, value in enumerate(df_matches.columns):
                worksheet.write(0, col_num, value, fmt_header)
    output.seek(0)
    return output

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="TPRM Vendor Comparison", layout="wide")
st.title("🛡️ TPRM Vendor Comparison Tool (Web Version)")
st.write("Upload Blue Book and Export Data, choose matching options, review and export results.")

# Upload files
with st.sidebar:
    st.header("Upload files")
    bb_file = st.file_uploader("📘 Blue Book (Excel or CSV)", type=["xlsx", "xls", "csv"])
    exp_file = st.file_uploader("📊 Export Data (Excel or CSV)", type=["xlsx", "xls", "csv"])

    st.markdown("---")
    st.header("Matching options")
    method = st.selectbox("Matching method", options=[
        ("Exact normalized match", "exact"),
        ("Fuzzy (RapidFuzz token sort)", "rapidfuzz") if HAS_RAPIDFUZZ else ("Fuzzy (difflib fallback)", "difflib"),
        ("Fuzzy (difflib)", "difflib")
    ], format_func=lambda x: x[0])[1]

    threshold = st.slider("Minimum match score to consider as MATCH", min_value=0, max_value=100, value=85, step=1)
    max_preview = st.slider("Max preview rows", min_value=5, max_value=50, value=10)
    st.markdown("---")
    st.write("RapidFuzz available:" , "✅" if HAS_RAPIDFUZZ else "❌ (using difflib)")

# Main area
if bb_file and exp_file:
    bb_df = load_file(bb_file)
    exp_df = load_file(exp_file)

    st.success(f"Blue Book loaded: {len(bb_df)} rows")
    st.success(f"Export Data loaded: {len(exp_df)} rows")

    # Column selection
    st.subheader("Step 2 — Select columns to compare")
    col1, col2 = st.columns(2)
    bb_col = col1.selectbox("Blue Book column", options=list(bb_df.columns))
    exp_col = col2.selectbox("Export Data column", options=list(exp_df.columns))

    # Preview
    st.subheader("Preview")
    p1, p2 = st.columns(2)
    p1.write(bb_df[bb_col].head(max_preview))
    p2.write(exp_df[exp_col].head(max_preview))

    # Run matching
    if st.button("🔍 Compare Vendors"):
        st.subheader("Matching in progress")
        progress = st.progress(0)
        total = len(exp_df)
        results = []
        # Precompute normalized choices from Blue Book
        bb_norm = bb_df[bb_col].astype(str).apply(normalize_text).tolist()
        # Keep mapping from normalized -> original (first occurrence)
        bb_map = {}
        for orig in bb_df[bb_col].astype(str).tolist():
            n = normalize_text(orig)
            if n not in bb_map:
                bb_map[n] = orig

        choices = bb_norm

        for i, val in enumerate(exp_df[exp_col].astype(str).tolist()):
            norm_val = normalize_text(val)
            if method == 'exact':
                match_norm, score = best_match(norm_val, set(choices), method='exact')
            elif method == 'rapidfuzz' and HAS_RAPIDFUZZ:
                match_norm, score = best_match(norm_val, choices, method='rapidfuzz')
            else:
                match_norm, score = best_match(norm_val, choices, method='difflib')

            matched_original = bb_map.get(match_norm, "") if match_norm else ""
            status = "Matched" if score >= threshold and match_norm else "Missing"
            results.append({
                "export_value_original": val,
                "export_value_norm": norm_val,
                "best_match_norm": match_norm or "",
                "best_match_original": matched_original,
                "score": score,
                "status": status
            })
            if i % 10 == 0:
                progress.progress(int((i / total) * 100))
        progress.progress(100)
        st.success("Matching complete")

        results_df = pd.DataFrame(results)
        # Summary metrics
        matched_count = (results_df['status'] == 'Matched').sum()
        missing_count = (results_df['status'] == 'Missing').sum()
        match_rate = round((matched_count / len(results_df)) * 100, 2) if len(results_df) else 0

        mcol1, mcol2, mcol3 = st.columns(3)
        mcol1.metric("Matched", matched_count)
        mcol2.metric("Missing", missing_count)
        mcol3.metric("Match Rate", f"{match_rate}%")

        st.subheader("Results table (editable)")
        # Allow user to edit the best match or status
        editable = st.experimental_data_editor(results_df[[
            "export_value_original", "best_match_original", "score", "status"
        ]], num_rows="dynamic")

        # Prepare final DataFrames for export
        final_df = editable.rename(columns={
            "export_value_original": "Export Value",
            "best_match_original": "Matched Blue Book Value",
            "score": "Score",
            "status": "Status"
        })

        missing_df = final_df[final_df['Status'] != 'Matched'].copy()
        matches_df = final_df[final_df['Status'] == 'Matched'].copy()

        st.subheader("Missing Vendors")
        st.dataframe(missing_df)

        # Downloads
        csv_bytes = missing_df.to_csv(index=False).encode('utf-8')
        st.download_button("📄 Download Missing CSV", csv_bytes, file_name="missing_vendors.csv", mime="text/csv")

        excel_bytes = create_excel_bytes(missing_df, matches_df)
        st.download_button("📥 Download Excel (Matches + Missing)", excel_bytes, file_name=f"vendor_compare_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.info("Tip: edit the 'Matched Blue Book Value' or 'Status' directly in the table to correct false positives/negatives, then re-download.")
else:
    st.info("Upload both Blue Book and Export Data files in the sidebar to begin.")
