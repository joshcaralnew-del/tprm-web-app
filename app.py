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
# Preprocessing helpers
# ---------------------------
COMPANY_SUFFIXES = [
    r'\binc\b', r'\bincorporated\b', r'\bcorp\b', r'\bcorporation\b',
    r'\bllc\b', r'\blimited\b', r'\bltd\b', r'\bco\b', r'\bcompany\b',
    r'\bplc\b', r'\bllp\b', r'\bpc\b', r'\bgroup\b', r'\bservices\b'
]

STOPWORDS = set([
    'the', 'and', 'of', 'for', 'a', 'an', 'in', 'on', 'at', 'by', 'to', 'with', 'solutions', 'systems'
])

def normalize_text(s):
    """Lowercase, strip, remove punctuation, collapse whitespace."""
    if pd.isna(s):
        return ""
    s = str(s)
    s = s.strip().lower()
    # remove common company suffixes
    for suf in COMPANY_SUFFIXES:
        s = re.sub(suf, '', s)
    # remove punctuation except spaces
    s = re.sub(r'[^\w\s]', ' ', s)
    # collapse whitespace
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def tokenize(s):
    """Return list of tokens excluding stopwords."""
    s = normalize_text(s)
    tokens = [t for t in s.split() if t and t not in STOPWORDS]
    return tokens

def jaccard_score(a_tokens, b_tokens):
    if not a_tokens or not b_tokens:
        return 0.0
    a_set = set(a_tokens)
    b_set = set(b_tokens)
    inter = a_set.intersection(b_set)
    union = a_set.union(b_set)
    return len(inter) / len(union) if union else 0.0

# ---------------------------
# Matching helpers
# ---------------------------
def rapidfuzz_top_matches(query, choices, top_n=3):
    """Return list of (choice, score_dict) using multiple rapidfuzz scorers."""
    # use extract with token_sort_ratio to get candidates quickly
    candidates = rf_process.extract(query, choices, scorer=rf_fuzz.token_sort_ratio, limit=top_n*3)
    # candidates: list of (choice, score, idx)
    results = []
    for cand in candidates[:top_n]:
        choice = cand[0]
        scores = {
            'token_sort_ratio': int(rf_fuzz.token_sort_ratio(query, choice)),
            'token_set_ratio': int(rf_fuzz.token_set_ratio(query, choice)),
            'partial_ratio': int(rf_fuzz.partial_ratio(query, choice)),
            'ratio': int(rf_fuzz.ratio(query, choice))
        }
        results.append((choice, scores))
    return results

def difflib_top_matches(query, choices, top_n=3):
    """Fallback: compute top_n by SequenceMatcher ratio."""
    scored = []
    for c in choices:
        score = int(difflib.SequenceMatcher(None, query, c).ratio() * 100)
        scored.append((c, score))
    scored.sort(key=lambda x: x[1], reverse=True)
    results = []
    for c, s in scored[:top_n]:
        # approximate breakdown: use same score for all fields
        scores = {
            'token_sort_ratio': s,
            'token_set_ratio': s,
            'partial_ratio': s,
            'ratio': s
        }
        results.append((c, scores))
    return results

def get_top_matches(query_raw, choices_norm, top_n=3):
    """Return top_n candidates with score dicts and jaccard."""
    q_norm = normalize_text(query_raw)
    q_tokens = tokenize(query_raw)
    if not q_norm:
        return []
    if HAS_RAPIDFUZZ:
        raw_matches = rapidfuzz_top_matches(q_norm, choices_norm, top_n=top_n)
    else:
        raw_matches = difflib_top_matches(q_norm, choices_norm, top_n=top_n)
    enriched = []
    for choice_norm, scores in raw_matches:
        choice_tokens = tokenize(choice_norm)
        jacc = jaccard_score(q_tokens, choice_tokens)
        scores['jaccard'] = round(jacc * 100, 2)
        enriched.append((choice_norm, scores))
    return enriched

def composite_score(scores, weights):
    """Compute weighted composite from score dict and weights dict."""
    # scores keys: token_sort_ratio, token_set_ratio, partial_ratio, ratio, jaccard
    total = 0.0
    wsum = 0.0
    for k, w in weights.items():
        val = scores.get(k, 0)
        total += w * val
        wsum += w
    return (total / wsum) if wsum else 0.0

# ---------------------------
# File & export helpers
# ---------------------------
def load_file(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def create_excel_bytes(df_missing, df_matches):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_matches.empty:
            df_matches.to_excel(writer, sheet_name='Matches', index=False)
        else:
            pd.DataFrame(columns=["Export Value", "Matched Blue Book Value", "Composite Score", "Status"]).to_excel(writer, sheet_name='Matches', index=False)
        if not df_missing.empty:
            df_missing.to_excel(writer, sheet_name='Missing', index=False)
        else:
            pd.DataFrame(columns=["Export Value", "Matched Blue Book Value", "Composite Score", "Status"]).to_excel(writer, sheet_name='Missing', index=False)
        workbook = writer.book
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white'})
        # apply header format
        for sheet in ['Matches', 'Missing']:
            worksheet = writer.sheets[sheet]
            worksheet.set_column(0, 10, 30)
            # write header formatting
            for col_num, header in enumerate(writer.sheets[sheet].get_default_row_height() and [] or []):
                pass
    output.seek(0)
    return output

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="TPRM Vendor Comparison (Enhanced)", layout="wide")
st.title("🛡️ TPRM Vendor Comparison Tool — Enhanced Matching")
st.write("Upload Blue Book and Export Data, tune matching weights, inspect top candidates, and export results.")

# Sidebar: uploads and matching controls
with st.sidebar:
    st.header("Files")
    bb_file = st.file_uploader("📘 Blue Book (Excel or CSV)", type=["xlsx", "xls", "csv"])
    exp_file = st.file_uploader("📊 Export Data (Excel or CSV)", type=["xlsx", "xls", "csv"])
    st.markdown("---")
    st.header("Matching controls")
    st.write("Adjust how different scorers contribute to the final composite score.")
    w_token_sort = st.slider("Weight: token_sort_ratio", 0.0, 2.0, 1.0, 0.1)
    w_token_set = st.slider("Weight: token_set_ratio", 0.0, 2.0, 1.0, 0.1)
    w_partial = st.slider("Weight: partial_ratio", 0.0, 2.0, 0.6, 0.1)
    w_ratio = st.slider("Weight: ratio", 0.0, 2.0, 0.4, 0.1)
    w_jaccard = st.slider("Weight: jaccard (token overlap)", 0.0, 2.0, 1.0, 0.1)
    threshold = st.slider("Composite score threshold to mark as MATCH", 0, 100, 80, 1)
    top_n = st.slider("Top N candidates to return", 1, 5, 3, 1)
    st.markdown("---")
    st.write("RapidFuzz available:", "✅" if HAS_RAPIDFUZZ else "❌ (difflib fallback)")
    st.markdown("Tips: raise jaccard to favor token overlap; raise partial_ratio to favor substring matches.")

# Main area
if bb_file and exp_file:
    try:
        bb_df = load_file(bb_file)
        exp_df = load_file(exp_file)
    except Exception as e:
        st.error(f"Error reading files: {e}")
        st.stop()

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
    p1.write(bb_df[bb_col].head(10))
    p2.write(exp_df[exp_col].head(10))

    # Prepare normalized choices list and mapping to original
    bb_norm_list = bb_df[bb_col].astype(str).apply(normalize_text).tolist()
    bb_map = {}
    for orig in bb_df[bb_col].astype(str).tolist():
        n = normalize_text(orig)
        if n not in bb_map:
            bb_map[n] = orig

    # Matching run
    if st.button("🔍 Compare Vendors"):
        st.subheader("Matching in progress")
        progress = st.progress(0)
        total = len(exp_df)
        results = []
        weights = {
            'token_sort_ratio': w_token_sort,
            'token_set_ratio': w_token_set,
            'partial_ratio': w_partial,
            'ratio': w_ratio,
            'jaccard': w_jaccard
        }

        choices = bb_norm_list

        for i, val in enumerate(exp_df[exp_col].astype(str).tolist()):
            top_matches = get_top_matches(val, choices, top_n=top_n)
            row = {
                "Export Value": val,
                "Export Norm": normalize_text(val)
            }
            # add candidate columns dynamically
            for idx in range(top_n):
                if idx < len(top_matches):
                    cand_norm, score_dict = top_matches[idx]
                    comp = composite_score(score_dict, weights)
                    row.update({
                        f"Candidate {idx+1} Norm": cand_norm,
                        f"Candidate {idx+1} Orig": bb_map.get(cand_norm, ""),
                        f"Candidate {idx+1} TokenSort": score_dict.get('token_sort_ratio', 0),
                        f"Candidate {idx+1} TokenSet": score_dict.get('token_set_ratio', 0),
                        f"Candidate {idx+1} Partial": score_dict.get('partial_ratio', 0),
                        f"Candidate {idx+1} Ratio": score_dict.get('ratio', 0),
                        f"Candidate {idx+1} Jaccard": score_dict.get('jaccard', 0),
                        f"Candidate {idx+1} Composite": round(comp, 2)
                    })
                else:
                    # empty placeholders
                    row.update({
                        f"Candidate {idx+1} Norm": "",
                        f"Candidate {idx+1} Orig": "",
                        f"Candidate {idx+1} TokenSort": 0,
                        f"Candidate {idx+1} TokenSet": 0,
                        f"Candidate {idx+1} Partial": 0,
                        f"Candidate {idx+1} Ratio": 0,
                        f"Candidate {idx+1} Jaccard": 0,
                        f"Candidate {idx+1} Composite": 0.0
                    })
            # Decide default status using top candidate composite
            top_comp = row.get("Candidate 1 Composite", 0.0)
            row["Default Status"] = "Matched" if top_comp >= threshold else "Missing"
            results.append(row)

            # update progress
            if total:
                if (i + 1) % max(1, total // 100) == 0 or i == total - 1:
                    progress.progress(int(((i + 1) / total) * 100))
        progress.progress(100)
        st.success("Matching complete")

        results_df = pd.DataFrame(results)

        # Summary metrics using Default Status
        matched_count = (results_df['Default Status'] == 'Matched').sum()
        missing_count = (results_df['Default Status'] == 'Missing').sum()
        match_rate = round((matched_count / len(results_df)) * 100, 2) if len(results_df) else 0

        mcol1, mcol2, mcol3 = st.columns(3)
        mcol1.metric("Matched (default)", matched_count)
        mcol2.metric("Missing (default)", missing_count)
        mcol3.metric("Match Rate (default)", f"{match_rate}%")

        st.subheader("Results (top candidates shown)")
        st.write("You can edit the 'Selected Match' and 'Status' columns below before exporting.")

        # Build a simplified editable table: Export Value, Selected Match, Composite, Status
        editable_rows = []
        for _, r in results_df.iterrows():
            selected = r.get("Candidate 1 Orig", "")  # default to candidate 1 original
            comp = r.get("Candidate 1 Composite", 0.0)
            editable_rows.append({
                "Export Value": r["Export Value"],
                "Selected Match": selected,
                "Composite Score": comp,
                "Status": r["Default Status"]
            })
        editable_df = pd.DataFrame(editable_rows)

        # Compatibility wrapper for Streamlit data editor
        def show_editable_dataframe(df):
            try:
                edited = st.data_editor(df, num_rows="dynamic")
                return edited
            except Exception:
                pass
            try:
                edited = st.experimental_data_editor(df, num_rows="dynamic")
                return edited
            except Exception:
                pass
            st.warning("Interactive table editor not available in this Streamlit version. Showing read-only results.")
            st.dataframe(df)
            return df

        editable = show_editable_dataframe(editable_df)

        # Prepare final DataFrames for export
        final_df = editable.copy()
        # Ensure columns exist
        for c in ["Export Value", "Selected Match", "Composite Score", "Status"]:
            if c not in final_df.columns:
                final_df[c] = ""

        missing_df = final_df[final_df['Status'] != 'Matched'].copy()
        matches_df = final_df[final_df['Status'] == 'Matched'].copy()

        st.subheader("Missing Vendors (after edits)")
        st.dataframe(missing_df)

        # Downloads
        csv_bytes = missing_df.to_csv(index=False).encode('utf-8')
        st.download_button("📄 Download Missing CSV", csv_bytes, file_name="missing_vendors.csv", mime="text/csv")

        try:
            excel_bytes = create_excel_bytes(missing_df, matches_df)
            st.download_button(
                "📥 Download Excel (Matches + Missing)",
                excel_bytes,
                file_name=f"vendor_compare_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Excel export failed: {e}")

        st.info("Tip: adjust weights and threshold in the sidebar to improve automatic matching, then re-run Compare.")
else:
    st.info("Upload both Blue Book and Export Data files in the sidebar to begin.")
