import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel (Final)", layout="wide")
st.title("Refund Panel — FINAL VERSION (Special Rule + Rule1/2/3 + Universal Stray Join)")

# ---------------------------
# File Upload Section
# ---------------------------
uploaded = st.file_uploader("Upload INPUT file (CSV/XLSX): allotmentdetails", type=["csv","xlsx","xls"])
manual_file = st.file_uploader("OPTIONAL: Upload manual refund file (Total_Details.xlsx) to compare", type=["xlsx","xls","csv"])

if uploaded is None:
    st.info("Upload the INPUT allotment file to begin.")
    st.stop()

# Read input
fname = uploaded.name.lower()
try:
    if fname.endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)
except Exception as e:
    st.error(f"Error reading input file: {e}")
    st.stop()

st.subheader("Input Preview")
st.dataframe(df.head(50), use_container_width=True)

# ---------------------------
# Column Normalization
# ---------------------------
expected_cols = [
    "APPLNO","regfeepaid","forefit",
    "Allot_1","Allot_2","Allot_3",
    "JoinStatus_1","JoinStatus_2","JoinStatus_3",
    "JoinStray","Stray","Remarks"
]

for c in expected_cols:
    if c not in df.columns:
        # Check for common alternatives
        alternatives = [col for col in df.columns if col.lower().replace(" ", "") == c.lower()]
        if alternatives:
            df.rename(columns={alternatives[0]: c}, inplace=True)
        else:
            df[c] = ""

# Numeric cleanup
def to_num(df_local, col):
    if col in df_local.columns:
        df_local[col] = df_local[col].astype(str).str.replace(",", "").replace("nan", "0").replace("None", "0")
        df_local[col] = pd.to_numeric(df_local[col], errors="coerce").fillna(0)
    else:
        df_local[col] = 0

to_num(df, "regfeepaid")
to_num(df, "forefit")

# Helper functions
def sval(x):
    return "" if pd.isna(x) else str(x).strip()

def sstatus(x):
    return sval(x).upper()


# ============================================================
#                   REFUND COMPUTATION ENGINE
# ============================================================
def compute_refund(row):

    reg_paid = float(row.get("regfeepaid", 0) or 0)
    forefit_old = float(row.get("forefit", 0) or 0)

    js1 = sstatus(row.get("JoinStatus_1", ""))
    js2 = sstatus(row.get("JoinStatus_2", ""))
    js3 = sstatus(row.get("JoinStatus_3", ""))
    join_stray = sstatus(row.get("JoinStray", ""))

    allot1 = sval(row.get("Allot_1", ""))
    allot2 = sval(row.get("Allot_2", ""))
    allot3 = sval(row.get("Allot_3", ""))
    straycol = sval(row.get("Stray", ""))
    remarks = sval(row.get("Remarks", ""))

    # --------------------------------------------------------
    # SPECIAL RULE (Top Priority)
    # --------------------------------------------------------
    if (js1 == "Y" or js2 == "Y") and js3 == "TC" and join_stray == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "SPECIAL: Joined P1/P2 + TC + Stray → FULL REFUND"
        })

    # --------------------------------------------------------
    # UNIVERSAL STRAY RULE (Your final instruction)
    # JoinStray = Y ALWAYS means JOINED
    # --------------------------------------------------------
    if join_stray == "Y":

        # CASE A: Paid twice (forefit > 0 & regfeepaid > 0)
        if forefit_old > 0 and reg_paid > 0:
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": forefit_old,
                "RefundCategory": "Stray Join → Refund SECOND payment"
            })

        # CASE B: Paid once → full refund
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": 0,
            "RefundCategory": "Stray Join → FULL REFUND"
        })

    # --------------------------------------------------------
    # RULE 2: No allotment anywhere + NOT stray (since stray handled above)
    # --------------------------------------------------------
    if allot1 == "" and allot2 == "" and allot3 == "" and straycol == "":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "No allotment → FULL REFUND"
        })

    # --------------------------------------------------------
    # RULE 1 & RULE 3 now simplified thanks to universal stray rule
    # (Already handled above)
    # --------------------------------------------------------

    # --------------------------------------------------------
    # Joined in P1 or P2
    # --------------------------------------------------------
    if js1 == "Y" or js2 == "Y":
        if "TC" in (js1, js2, js3):
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": forefit_old + reg_paid,
                "RefundCategory": "Joined then TC → NO REFUND"
            })
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "Joined P1/P2 → FULL REFUND"
        })

    # --------------------------------------------------------
    # Phase 3 join (not stray)
    # --------------------------------------------------------
    if js1 == "" and js2 == "" and js3 == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": forefit_old,
            "RefundCategory": "Joined P3 → FULL REFUND"
        })

    # --------------------------------------------------------
    # Not joined anywhere → forfeit
    # --------------------------------------------------------
    if js1 in ("N", "TC") and js2 in ("N", "TC"):
        if js3 == "N":
            # no join at all
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": forefit_old + reg_paid,
                "RefundCategory": "No Join in P1/P2/P3 → FORFEIT"
            })
        else:
            return pd.Series({
                "RefundAmount": 0,
                "NewForefit": forefit_old + reg_paid,
                "RefundCategory": "Not Joined P1/P2 → FORFEIT"
            })

    # --------------------------------------------------------
    # Fallback
    # --------------------------------------------------------
    return pd.Series({
        "RefundAmount": 0,
        "NewForefit": forefit_old,
        "RefundCategory": "CHECK MANUALLY"
    })


# ============================================================
# Apply Refund Logic
# ============================================================
st.subheader("Processing Refunds...")
calc = df.apply(compute_refund, axis=1)
df_out = pd.concat([df, calc], axis=1)

st.subheader("Output Preview")
st.dataframe(df_out.head(100), use_container_width=True)

# Summary
col1, col2, col3 = st.columns(3)
col1.metric("Total regfeepaid", f"{df_out['regfeepaid'].sum():,.2f}")
col2.metric("Total Refund", f"{df_out['RefundAmount'].sum():,.2f}")
col3.metric("Total Forfeit", f"{df_out['NewForefit'].sum():,.2f}")


# ============================================================
# Optional Manual Comparison
# ============================================================
if manual_file is not None:

    st.subheader("Comparing with Manual Refund File…")

    mfname = manual_file.name.lower()
    try:
        if mfname.endswith(".csv"):
            df_manual = pd.read_csv(manual_file, dtype=str)
        else:
            df_manual = pd.read_excel(manual_file, dtype=str)
    except:
        st.error("Could not read manual file.")
        st.stop()

    if "APPLNO" not in df_manual.columns:
        st.error("Manual file must contain APPLNO column")
        st.stop()

    if "Refund" not in df_manual.columns:
        st.error("Manual file must contain Refund column")
        st.stop()

    df_manual["APPLNO"] = df_manual["APPLNO"].astype(str).str.strip()
    df_manual["Refund"] = pd.to_numeric(df_manual["Refund"], errors="coerce").fillna(0)

    df_cmp = df_out.merge(df_manual[["APPLNO","Refund"]].rename(columns={"Refund":"ManualRefund"}),
                          on="APPLNO", how="left")

    df_cmp["ManualRefund"] = df_cmp["ManualRefund"].fillna(0)
    df_cmp["Diff"] = df_cmp["RefundAmount"] - df_cmp["ManualRefund"]
    df_cmp["Mismatch"] = df_cmp["Diff"].abs() > 0.0001

    mismatches = df_cmp[df_cmp["Mismatch"]]

    st.metric("Mismatches", len(mismatches))
    st.dataframe(mismatches.head(50), use_container_width=True)


# ============================================================
# EXPORT SECTION
# ============================================================
st.subheader("Download Results")

csv_bytes = df_out.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv_bytes, "refund_output.csv", mime="text/csv")


# Colored Excel
def colored_excel(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Refund")
            wb = writer.book
            ws = writer.sheets["Refund"]

            fmt_ref = wb.add_format({"bg_color":"#c6efce"})
            fmt_for = wb.add_format({"bg_color":"#ffc7ce"})
            fmt_chk = wb.add_format({"bg_color":"#ffeb9c"})

            for r in range(len(df)):
                cat = df.loc[r,"RefundCategory"].lower()
                if "refund" in cat:
                    ws.set_row(r+1, cell_format=fmt_ref)
                elif "forfeit" in cat:
                    ws.set_row(r+1, cell_format=fmt_for)
                else:
                    ws.set_row(r+1, cell_format=fmt_chk)

        return output.getvalue()
    except:
        return None

excel_bytes = colored_excel(df_out)
if excel_bytes:
    st.download_button("Download Colored Excel", excel_bytes, "refund_output_colored.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.warning("Excel coloring unavailable (xlsxwriter missing). CSV still available.")

st.success("Refund processing completed successfully!")
