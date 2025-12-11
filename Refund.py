import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel", layout="wide")
st.title("Refund Panel – Final Version (Previous Code + Universal Stray Rule)")

# -----------------------------
# FILE UPLOAD SECTION
# -----------------------------
uploaded = st.file_uploader("Upload allotment file (CSV or XLSX)", type=["csv","xlsx","xls"])
manual_file = st.file_uploader("Optional: Upload manual refund sheet (Total_Details.xlsx) for comparison", type=["xlsx","xls","csv"])

if uploaded is None:
    st.info("Upload the allotment file to continue.")
    st.stop()

# Read input file
try:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)
except:
    st.error("Error loading file.")
    st.stop()

# -----------------------------
# PREPARE COLUMNS
# -----------------------------
expected_cols = [
    "APPLNO","regfeepaid","forefit",
    "Allot_1","Allot_2","Allot_3",
    "JoinStatus_1","JoinStatus_2","JoinStatus_3",
    "JoinStray","Stray","Remarks"
]

for c in expected_cols:
    if c not in df.columns:
        df[c] = ""

def clean_num(df_local, col):
    if col not in df_local.columns:
        df_local[col] = 0
    df_local[col] = (
        df_local[col]
        .astype(str)
        .str.replace(",", "")
        .replace("nan", "0")
        .replace("None", "0")
    )
    df_local[col] = pd.to_numeric(df_local[col], errors="coerce").fillna(0)

clean_num(df, "regfeepaid")
clean_num(df, "forefit")

def sval(x):
    return "" if pd.isna(x) else str(x).strip()

def sstatus(x):
    return sval(x).upper()

# ============================================================
# REFUND ENGINE (Version A + Universal Stray Rule)
# ============================================================
def compute_refund(row):

    reg_paid = float(row["regfeepaid"])
    forefit_old = float(row["forefit"])

    js1 = sstatus(row["JoinStatus_1"])
    js2 = sstatus(row["JoinStatus_2"])
    js3 = sstatus(row["JoinStatus_3"])
    join_stray = sstatus(row["JoinStray"])

    allot1 = sval(row["Allot_1"])
    allot2 = sval(row["Allot_2"])
    allot3 = sval(row["Allot_3"])
    straycol = sval(row["Stray"])

    # ---------------------------------------------
    # SPECIAL RULE
    # ---------------------------------------------
    if (js1 == "Y" or js2 == "Y") and js3 == "TC" and join_stray == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "SPECIAL: P1/P2 Joined + P3 TC + Stray → Full Refund"
        })

    # ---------------------------------------------
    # UNIVERSAL STRAY JOIN RULE (your final instruction)
    # ---------------------------------------------
    if join_stray == "Y":

        # Paid twice → refund second payment only
        if forefit_old > 0 and reg_paid > 0:
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": forefit_old,
                "RefundCategory": "Stray Join → Refund SECOND payment"
            })

        # Paid once → full refund
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": 0,
            "RefundCategory": "Stray Join → Full Refund"
        })

    # ---------------------------------------------
    # RULE 2: No allotment anywhere → Refund
    # ---------------------------------------------
    if allot1 == "" and allot2 == "" and allot3 == "" and straycol == "":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "No Allotment → Full Refund"
        })

    # ---------------------------------------------
    # RULE: Phase 1 or Phase 2 Joined
    # ---------------------------------------------
    if js1 == "Y" or js2 == "Y":
        # If any TC later → no refund
        if js3 == "TC" or js1 == "TC" or js2 == "TC":
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": forefit_old + reg_paid,
                "RefundCategory": "Joined then TC → No Refund"
            })
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "Joined P1/P2 → Full Refund"
        })

    # ---------------------------------------------
    # RULE: Joined Phase 3
    # ---------------------------------------------
    if js1 == "" and js2 == "" and js3 == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "Joined P3 → Full Refund"
        })

    # ---------------------------------------------
    # RULE: Not joined anywhere → Forfeit
    # ---------------------------------------------
    if js1 in ("N","TC") and js2 in ("N","TC"):
        return pd.Series({
            "RefundAmount": 0,
            "NewForefit": forefit_old + reg_paid,
            "RefundCategory": "Not Joined → Forfeit"
        })

    # ---------------------------------------------
    # FALLBACK
    # ---------------------------------------------
    return pd.Series({
        "RefundAmount": 0,
        "NewForefit": forefit_old,
        "RefundCategory": "Check Manually"
    })

# ============================================================
# APPLY RULES
# ============================================================
st.subheader("Computing refunds…")
calc = df.apply(compute_refund, axis=1)
df_out = pd.concat([df, calc], axis=1)
st.dataframe(df_out.head(100), use_container_width=True)

# SUMMARY
c1, c2, c3 = st.columns(3)
c1.metric("Total regfeepaid", f"{df_out['regfeepaid'].sum():,.2f}")
c2.metric("Total Refund", f"{df_out['RefundAmount'].sum():,.2f}")
c3.metric("Total Forfeit", f"{df_out['NewForefit'].sum():,.2f}")

# ============================================================
# EXPORT OUTPUT
# ============================================================
st.subheader("Download Results")

csv_bytes = df_out.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv_bytes, "refund_output.csv", mime="text/csv")

# Colored Excel
def make_xlsx(df_local):
    try:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            df_local.to_excel(writer, index=False, sheet_name="Refund")
            wb = writer.book
            ws = writer.sheets["Refund"]

            fmt_ref = wb.add_format({"bg_color": "#c6efce"})
            fmt_for = wb.add_format({"bg_color": "#ffc7ce"})
            fmt_chk = wb.add_format({"bg_color": "#ffeb9c"})

            for r in range(len(df_local)):
                cat = df_local.loc[r,"RefundCategory"].lower()
                if "refund" in cat:
                    ws.set_row(r+1, cell_format=fmt_ref)
                elif "forfeit" in cat:
                    ws.set_row(r+1, cell_format=fmt_for)
                else:
                    ws.set_row(r+1, cell_format=fmt_chk)

        return out.getvalue()
    except:
        return None

xlsx_bytes = make_xlsx(df_out)
if xlsx_bytes:
    st.download_button("Download Colored Excel", xlsx_bytes, "refund_output_colored.xlsx")

st.success("Processing completed successfully.")
