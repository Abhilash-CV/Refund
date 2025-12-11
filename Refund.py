import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel", layout="wide")

st.markdown("<h2 style='text-align:center;'>Refund Panel</h2>", unsafe_allow_html=True)
st.write("")

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader(
    "Upload Allotment File (CSV / Excel)", 
    type=["csv", "xlsx", "xls"]
)

if uploaded_file is None:
    st.info("Please upload the allotment file to begin.")
    st.stop()

# Detect file type
file_name = uploaded_file.name.lower()
try:
    if file_name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

st.subheader("Input Data")
st.dataframe(df, use_container_width=True)

# =========================
# REQUIRED COLUMNS
# =========================
required_cols = [
    "regfeepaid", "forefit",
    "Allot_1", "Allot_2", "Allot_3",
    "JoinStatus_1", "JoinStatus_2", "JoinStatus_3",
    "JoinStray", "Stray"
]

missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"Missing Required Columns: {missing_cols}")
    st.stop()

# Clean numeric fields
df["regfeepaid"] = pd.to_numeric(df["regfeepaid"], errors="coerce").fillna(0)
df["forefit"]     = pd.to_numeric(df["forefit"],     errors="coerce").fillna(0)

# =========================
# HELPERS
# =========================
def sval(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def sstatus(x):
    return sval(x).upper()

# =========================
# MAIN REFUND LOGIC
# =========================
def compute_refund(row):

    reg_paid = row.get("regfeepaid", 0)
    existing_forefit = row.get("forefit", 0)

    # Status fields
    js1 = sstatus(row.get("JoinStatus_1", ""))
    js2 = sstatus(row.get("JoinStatus_2", ""))
    js3 = sstatus(row.get("JoinStatus_3", ""))
    join_stray = sstatus(row.get("JoinStray", ""))

    # Allotment info
    allot1 = sval(row.get("Allot_1", ""))
    allot2 = sval(row.get("Allot_2", ""))
    allot3 = sval(row.get("Allot_3", ""))
    stray  = sval(row.get("Stray", ""))

    # =============================================================
    # SPECIAL RULE (Highest Priority – Checked First)
    # Joined P1/P2 + TC in P3 + Stray = FULL REFUND
    # =============================================================
    if (js1 == "Y" or js2 == "Y") and js3 == "TC" and join_stray == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Special Rule: Joined earlier + TC in P3 + Stray → full refund"
        })

    # =============================================================
    # 1) NO ALLOTMENT ANYWHERE → FULL REFUND
    # =============================================================
    if allot1 == "" and allot2 == "" and allot3 == "" and stray == "":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "No allotment – full refund"
        })

    # =============================================================
    # 2) JOINED IN PHASE 1 OR 2
    # =============================================================
    if js1 == "Y" or js2 == "Y":

        # If TC (and special rule didn't trigger) → NO REFUND
        if "TC" in (js1, js2, js3):
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Joined then TC – no refund"
            })

        # Otherwise → FULL REFUND
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Joined in phase 1/2 – full refund"
        })

    # =============================================================
    # 3) PHASE 3 / STRAY LOGIC
    # =============================================================
    # Only valid if NOT joined in phase 1 or 2
    if js1 == "" and js2 == "":

        # Joined in Phase 3 → Full Refund
        if js3 == "Y":
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": existing_forefit,
                "RefundCategory": "Joined in phase 3 – full refund"
            })

        # Joined through stray (JoinStray = Y AND js3 = Y)
        if join_stray == "Y" and js3 == "Y":
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": existing_forefit,
                "RefundCategory": "Joined via stray – full refund"
            })

        # Stray allotted but NOT joined → NO REFUND
        if join_stray == "Y" and js3 != "Y":
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Stray allotted but NOT joined – no refund"
            })

    # =============================================================
    # 4) NOT JOINED / TC IN PHASE 1 & 2 → FORFEIT
    # =============================================================
    if js1 in ("N", "TC") and js2 in ("N", "TC"):

        if js3 == "N":
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Not joined in 1/2/3 – forfeit"
            })

        return pd.Series({
            "RefundAmount": 0,
            "NewForefit": existing_forefit + reg_paid,
            "RefundCategory": "Not joined in 1/2 – forfeit"
        })

    # =============================================================
    # 5) DEFAULT → MANUAL CHECK
    # =============================================================
    return pd.Series({
        "RefundAmount": 0,
        "NewForefit": existing_forefit,
        "RefundCategory": "Check manually"
    })

# Apply logic to all rows
df_result = df.join(df.apply(compute_refund, axis=1))

# =========================
# SUMMARY DASHBOARD
# =========================
st.subheader("Summary")

c1, c2, c3 = st.columns(3)
c1.metric("Total Reg Fee Paid", f"{df_result['regfeepaid'].sum():,.2f}")
c2.metric("Total Refund", f"{df_result['RefundAmount'].sum():,.2f}")
c3.metric("Total Forfeit", f"{df_result['NewForefit'].sum():,.2f}")

st.subheader("Refund Output")
st.dataframe(df_result, use_container_width=True)

# =========================
# DOWNLOAD CSV (ALWAYS WORKS)
# =========================
csv_data = df_result.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download Refund CSV",
    data=csv_data,
    file_name="refund_output.csv",
    mime="text/csv"
)
