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
uploaded_file = st.file_uploader("Upload Allotment File (CSV / Excel)", 
                                 type=["csv", "xlsx", "xls"])

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
    st.error(f"Missing required columns: {missing_cols}")
    st.stop()

# Ensure numeric fields are clean
df["regfeepaid"] = pd.to_numeric(df["regfeepaid"], errors="coerce").fillna(0)
df["forefit"] = pd.to_numeric(df["forefit"], errors="coerce").fillna(0)


# =========================
# HELPER FUNCTIONS
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

    # Extract status fields
    js1 = sstatus(row.get("JoinStatus_1", ""))
    js2 = sstatus(row.get("JoinStatus_2", ""))
    js3 = sstatus(row.get("JoinStatus_3", ""))
    join_stray = sstatus(row.get("JoinStray", ""))

    # Allotments
    allot1 = sval(row.get("Allot_1", ""))
    allot2 = sval(row.get("Allot_2", ""))
    allot3 = sval(row.get("Allot_3", ""))
    stray = sval(row.get("Stray", ""))

    refund = 0
    new_forefit = existing_forefit
    reason = "Check manually"

    # -------------------------------------------------
    # 1) NO ALLOTMENT ANYWHERE → FULL REFUND
    # -------------------------------------------------
    if allot1 == "" and allot2 == "" and allot3 == "" and stray == "":
        refund = reg_paid
        reason = "No allotment – full refund"
        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": existing_forefit,
            "RefundCategory": reason
        })

    # -------------------------------------------------
    # 2) JOINED IN PHASE 1 OR 2
    # -------------------------------------------------
    if js1 == "Y" or js2 == "Y":

        # If later TC → NO REFUND
        if "TC" in (js1, js2, js3):
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Joined then TC – no refund"
        else:
            refund = reg_paid
            reason = "Joined in phase 1/2 – full refund"

        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": new_forefit,
            "RefundCategory": reason
        })

    # -------------------------------------------------
    # 3) PHASE 3 / STRAY LOGIC
    # -------------------------------------------------
    # Candidate did NOT join in Ph1 & Ph2
    if js1 == "" and js2 == "":

        # Joined in Phase 3 → FULL REFUND
        if js3 == "Y":
            refund = reg_paid
            reason = "Joined in phase 3 – full refund"
            return pd.Series({
                "RefundAmount": refund,
                "NewForefit": existing_forefit,
                "RefundCategory": reason
            })

        # Joined through stray (JoinStray=Y AND actual join (js3='Y'))
        if join_stray == "Y" and js3 == "Y":
            refund = reg_paid
            reason = "Joined via stray – full refund"
            return pd.Series({
                "RefundAmount": refund,
                "NewForefit": existing_forefit,
                "RefundCategory": reason
            })

        # Stray allotted but candidate DID NOT JOIN → NO REFUND
        if join_stray == "Y" and js3 != "Y":
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Stray allotted but NOT joined – no refund"
            return pd.Series({
                "RefundAmount": refund,
                "NewForefit": new_forefit,
                "RefundCategory": reason
            })

    # -------------------------------------------------
    # 4) NOT JOINED / TC IN PHASE 1 & 2 → FORFEIT
    # -------------------------------------------------
    if js1 in ("N", "TC") and js2 in ("N", "TC"):

        # If also not joined in Phase 3 → forfeit
        if js3 == "N":
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Not joined in 1/2/3 – forfeit"

        else:
            refund = 0
            new_forefit = existing_forefit + reg_paid
            reason = "Not joined in 1/2 – forfeit"

        return pd.Series({
            "RefundAmount": refund,
            "NewForefit": new_forefit,
            "RefundCategory": reason
        })

    # -------------------------------------------------
    # 5) ANYTHING ELSE → MANUAL CHECK
    # -------------------------------------------------
    return pd.Series({
        "RefundAmount": 0,
        "NewForefit": existing_forefit,
        "RefundCategory": "Check manually"
    })


# Apply the logic
df_result = df.join(df.apply(compute_refund, axis=1))


# =========================
# SUMMARY
# =========================
st.subheader("Summary")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total Reg Fee Paid", f"{df_result['regfeepaid'].sum():,.2f}")
with col2:
    st.metric("Total Refund", f"{df_result['RefundAmount'].sum():,.2f}")
with col3:
    st.metric("Total Forfeit", f"{df_result['NewForefit'].sum():,.2f}")

st.write("")
st.subheader("Refund Output")
st.dataframe(df_result, use_container_width=True)


# =========================
# DOWNLOAD AS CSV (WORKS ON ALL STREAMLIT SERVERS)
# =========================
csv_output = df_result.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download Refund CSV",
    data=csv_output,
    file_name="refund_output.csv",
    mime="text/csv"
)
