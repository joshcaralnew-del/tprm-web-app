import streamlit as st
import pandas as pd
import re

# ---------------------------
# Utility functions
# ---------------------------
def normalize_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    s = re.sub(r'[\s\-\_]+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s

def best_match(query, choices):
    q = normalize_text(query)
    if not q:
        return None, 0

    # Simple difflib fallback
    import difflib
    matches = difflib.get_close_matches(q, choices, n=1, cutoff=0.0)
    if matches:
        score = int(difflib.SequenceMatcher(None, q, matches[0]).ratio() * 100)
        return matches[0], score
    return None, 0

# ---------------------------
# Streamlit UI
# ---------------------------
st.title("🛡️ TPRM Vendor Comparison Tool (Web Version)")
st.write("Upload your Blue Book and Export Data files, select columns, and compare.")

# ---------------------------
# File Upload
# ---------------------------
bb_file = st.file_uploader("📘 Upload Blue Book (Excel or CSV)", type=["xlsx", "xls", "csv"])
exp_file = st.file_uploader("📊 Upload Export Data (Excel or CSV)", type=["xlsx", "xls", "csv"])

def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

if bb_file and exp_file:
    bb_df = load_file(bb_file)
    exp_df = load_file(exp_file)

    st.success(f"Blue Book loaded: {len(bb_df)} rows")
    st.success(f"Export Data loaded: {len(exp_df)} rows")

    # ---------------------------
    # Column Selection
    # ---------------------------
    st.subheader("Step 2️⃣ Select Columns to Compare")

    bb_col = st.selectbox("Blue Book Column", bb_df.columns)
    exp_col = st.selectbox("Export Data Column", exp_df.columns)

    # Preview
    st.write("### 📋 Preview")
    col1, col2 = st.columns(2)
    col1.write(bb_df[bb_col].head())
    col2.write(exp_df[exp_col].head())

    # ---------------------------
    # Compare
    # ---------------------------
    if st.button("🔍 Compare Vendors"):
        st.subheader("Results")

        bb_values = bb_df[bb_col].astype(str).apply(normalize_text).tolist()
        exp_values = exp_df[exp_col].astype(str).apply(normalize_text).tolist()

        bb_set = set(bb_values)
        exp_set = set(exp_values)

        matched = exp_set.intersection(bb_set)
        missing = exp_set - bb_set

        st.metric("Matched", len(matched))
        st.metric("Missing", len(missing))
        match_rate = round((len(matched) / len(exp_set)) * 100, 2)
        st.metric("Match Rate", f"{match_rate}%")

        st.write("### Missing Vendors")
        missing_list = sorted(list(missing))
        st.write(missing_list)

        # Download CSV
        missing_df = pd.DataFrame({"Missing Vendors": missing_list})
        st.download_button(
            "📄 Download Missing Vendors CSV",
            missing_df.to_csv(index=False),
            "missing_vendors.csv",
            "text/csv"
        )
