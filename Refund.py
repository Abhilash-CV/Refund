import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel", layout="wide")

st.markdown("<h2 style='text-align:center;'>Refund Panel</h2>", unsafe_allow_html=True)
st.write("")

# =========================
# File Upload
# =========================
uploaded_file = st.file_uploader("Upload Allotment File (CSV / Excel)", type=["csv", "xlsx", "xls"])

if uploaded_file is None:
    st.info("Please upload the allotment file to begin.")
    st.stop()

# Detect input format
name = uploaded_file.name.lower()
try:
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

st.subheader("Input Data")
st.dataframe(df, use_container_width=True)

# =========================
# Required Columns
# =========================
required_cols = [
    "regfeepaid", "forefit",
    "Allot_1", "Allot_2", "Allot_3",
    "JoinStatus_1", "JoinStatus_2", "JoinStatus_3",
    "JoinStray", "Stray"
]

missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

# Cleanup numeric fields
df["regfeepaid"] = pd.to_numeric(df["regfeepaid"], errors="coerce").fillna(0)
df["forefit"] = pd.to_numeric(df["forefit"], errors="coerce").fillna(0)

# Helper functions
def sval(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def sstatus(x):
    return sval(x).upper()


# =========================
# REFUND CALCULATION LOGIC
# =========================
def compute_refund(row):
    reg_paid = row.get("regfeepaid", 0) or 0
    existing_forefit = row.get("forefit", 0) or 0

    # Statuses
    js1 = sstatus(row.get("JoinStatus_1", ""))
    js2 = sstatus(row.get("JoinStatus_2", ""))
    js3 = sstatus(row.get("JoinStatus_3", ""))
    join_stray = sstatus(row.get("JoinStray", ""))

    # Allotments
    allot1 = sval(row.get("Allot_1", ""))
    allot2 = sval(row.get("Allot_2", ""))
    allot3 = sval(row.get("Allot_3", ""))
    stray = sval(row.get("Stray", ""))

    refund = 0.0
    new_forefit = existing_forefit
    reason = "Check manually"

    # ----------------------------------------------------------------------
    # 1) No allotment anywhere → FULL REFUND
    # ----------------------------------------------------------------------
    if allot1 == "" and allot2 == "" and allot3 == "" and stray == "":
        refund = reg_paid
        reason = "No allotment – full refund"
        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": existing_forefit,
            "RefundCategory": reason
        })

    # ----------------------------------------------------------------------
    # 2) Joined in Phase 1 or 2 → FULL REFUND unless later TC
    # ----------------------------------------------------------------------
    if js1 == "Y" or js2 == "Y":
        # If they later took TC → NO REFUND
        if "TC" in (js1, js2, js3):
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Joined then TC → No refund"
        else:
            refund = reg_paid
            reason = "Joined in phase 1/2 – full refund"

        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": new_forefit,
            "RefundCategory": reason
        })

    # ----------------------------------------------------------------------
    # 3) Not joined in 1 & 2, but joined in Phase 3 or Stray → FULL REFUND
    # ----------------------------------------------------------------------
    if js1 == "" and js2 == "" and (js3 == "Y" or join_stray == "Y"):
        refund = reg_paid
        reason = "Joined in phase 3 / Stray – full refund"
        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": existing_forefit,
            "RefundCategory": reason
        })

    # ----------------------------------------------------------------------
    # 4) Not joined / TC in Phase 1 & 2 → NO REFUND
    # ----------------------------------------------------------------------
    cond_not_join_1_2 = js1 in ("N", "TC") and js2 in ("N", "TC")

    if cond_not_join_1_2:
        # If also Not Joined in Phase 3 → NEW PAYMENT ALSO FORFEIT
        if js3 == "N":
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Not joined in 1/2/3 → Forfeit"
        else:
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Not joined in phase 1/2 → Forfeit"

        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": new_forefit,
            "RefundCategory": reason
        })

    # ----------------------------------------------------------------------
    # Otherwise unclassified
    # ----------------------------------------------------------------------
    return pd.Series({
        "RefundAmount": 0,
        "NewForefit": existing_forefit,
        "RefundCategory": "Check manually"
    })


# Apply refund logic
calc = df.apply(compute_refund, axis=1)
df_result = df.join(calc)

# =========================
# SUMMARY
# =========================
st.subheader("Summary")

c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Total Reg Fee Paid", f"{df_result['regfeepaid'].sum():,.2f}")
with c2:
    st.metric("Total Refund", f"{df_result['RefundAmount'].sum():,.2f}")
with c3:
    st.metric("Total Forfeit", f"{df_result['NewForefit'].sum():,.2f}")

st.write("")
st.subheader("Refund Output")
st.dataframe(df_result, use_container_width=True)

# =========================
# DOWNLOAD AS EXCEL
# =========================
def to_excel_bytes(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Refund")
    return output.getvalue()

excel_bytes = to_excel_bytes(df_result)

st.download_button(
    label="Download Refund Excel",
    data=excel_bytes,
    file_name="refund_output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
