import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel", layout="wide")

st.markdown("<h2 style='text-align:center;'>Refund Panel</h2>", unsafe_allow_html=True)
st.write("")

# =========================
# File upload
# =========================
uploaded_file = st.file_uploader("Upload Allotment File (CSV / Excel)", type=["csv", "xlsx", "xls"])

if uploaded_file is None:
    st.info("Please upload the allotment file to start.")
    st.stop()

# Detect file type
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
# Basic validation
# =========================
required_cols = [
    "regfeepaid",
    "forefit",
    "Allot_1", "Allot_2", "Allot_3",
    "JoinStatus_1", "JoinStatus_2", "JoinStatus_3",
    "JoinStray",
    "Stray"
]

missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

# Ensure numeric
df["regfeepaid"] = pd.to_numeric(df["regfeepaid"], errors="coerce").fillna(0)
df["forefit"] = pd.to_numeric(df["forefit"], errors="coerce").fillna(0)

def s_val(x):
    """Safe string for status/allotment checks."""
    if pd.isna(x):
        return ""
    return str(x).strip()

def s_status(x):
    """Status in upper-case."""
    return s_val(x).upper()

# =========================
# Refund computation logic
# =========================
def compute_refund(row):
    reg_paid = row.get("regfeepaid", 0) or 0
    existing_forefit = row.get("forefit", 0) or 0

    # Status fields
    js1 = s_status(row.get("JoinStatus_1", ""))
    js2 = s_status(row.get("JoinStatus_2", ""))
    js3 = s_status(row.get("JoinStatus_3", ""))
    join_stray = s_status(row.get("JoinStray", ""))

    # Allotments
    allot1 = s_val(row.get("Allot_1", ""))
    allot2 = s_val(row.get("Allot_2", ""))
    allot3 = s_val(row.get("Allot_3", ""))
    stray = s_val(row.get("Stray", ""))

    refund = 0.0
    new_forefit = existing_forefit
    reason = "Check manually"

    # 1) No allotment in any phase/stray -> full refund
    if allot1 == "" and allot2 == "" and allot3 == "" and stray == "":
        refund = reg_paid
        new_forefit = existing_forefit
        reason = "No allotment in any phase"

    # 2) Joined in phase 1 or 2 -> full refund
    elif js1 == "Y" or js2 == "Y":
        refund = reg_paid
        new_forefit = existing_forefit
        reason = "Joined in phase 1/2 (Y)"

    # 3) Only joined in phase 3 or via stray -> full refund
    elif js1 == "" and js2 == "" and (js3 == "Y" or join_stray == "Y"):
        refund = reg_paid
        new_forefit = existing_forefit
        reason = "Joined in phase 3 / Stray"

    # 4) Not joined / TC in phase 1 & 2 -> no refund, forfeit reg fee
    #    Also handle case of paying again for 3rd round and still not joining
    else:
        cond_not_join_1_2 = (js1 in ("N", "TC") and js2 in ("N", "TC"))
        cond_not_join_3 = (js3 == "N" and join_stray != "Y")

        if cond_not_join_1_2 or cond_not_join_3:
            refund = 0.0
            # Move current payment to forefit
            new_forefit = existing_forefit + reg_paid
            reason = "Not joined / TC -> Forfeit"

    return pd.Series({
        "RefundAmount": refund,
        "NewForefit": new_forefit,
        "RefundCategory": reason
    })


st.subheader("Refund Calculation")
calc = df.apply(compute_refund, axis=1)
df_result = df.copy()
df_result = df_result.join(calc)

# =========================
# Summary
# =========================
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Total Reg Fee Paid", f"{df_result['regfeepaid'].sum():,.2f}")
with c2:
    st.metric("Total Refund", f"{df_result['RefundAmount'].sum():,.2f}")
with c3:
    additional_forefit = (df_result["NewForefit"] - df_result["forefit"]).sum()
    st.metric("Additional Forfeit", f"{additional_forefit:,.2f}")

st.write("")
st.dataframe(df_result, use_container_width=True)

# =========================
# Download processed file
# =========================
def to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Refund")
    return output.getvalue()

excel_bytes = to_excel_bytes(df_result)

st.download_button(
    label="Download Refund File (Excel)",
    data=excel_bytes,
    file_name="refund_output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
