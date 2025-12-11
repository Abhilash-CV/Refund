# Refund.py
import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Refund Panel (Final)", layout="wide")
st.title("Refund Panel — Final (Special + Rule1/2/3)")

# ----------------------
# Upload input & optional manual file
# ----------------------
uploaded = st.file_uploader("Upload input allotment file (CSV / XLSX). Use your input file (e.g. allotmentdetails.csv).", type=["csv","xlsx","xls"])
manual_file = st.file_uploader("(Optional) Upload your manually-correct file (Total_Details.xlsx) to compare", type=["xlsx","xls","csv"])

if uploaded is None:
    st.info("Upload an input file to run processing.")
    st.stop()

# read input
fname = uploaded.name.lower()
try:
    if fname.endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)
except Exception as e:
    st.error(f"Error reading input file: {e}")
    st.stop()

st.subheader("Input preview")
st.dataframe(df.head(100), use_container_width=True)

# Normalize expected columns (create if missing)
expected_cols = ["APPLNO","regfeepaid","forefit",
                 "Allot_1","Allot_2","Allot_3",
                 "JoinStatus_1","JoinStatus_2","JoinStatus_3",
                 "JoinStray","Stray","Remarks"]
for c in expected_cols:
    if c not in df.columns:
        # try common variants
        for alt in ["ApplNo","Applno","Appl_No","Appl No"]:
            if alt in df.columns:
                df.rename(columns={alt:c}, inplace=True)
                break
        else:
            df[c] = ""

# normalise numeric columns
def to_num_col(df_local, col):
    if col in df_local.columns:
        df_local[col] = df_local[col].astype(str).str.replace(",","").replace("nan","0").replace("None","0")
        df_local[col] = pd.to_numeric(df_local[col], errors="coerce").fillna(0)
    else:
        df_local[col] = 0

to_num_col(df, "regfeepaid")
to_num_col(df, "forefit")

# helper functions
def sval(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def sstatus(x):
    return sval(x).upper()

# ----------------------
# Refund rules implementation (Special -> Rule1/2/3 -> Standard)
# ----------------------
def compute_refund(row):
    reg_paid = float(row.get("regfeepaid", 0) or 0)
    existing_forefit = float(row.get("forefit", 0) or 0)

    js1 = sstatus(row.get("JoinStatus_1",""))
    js2 = sstatus(row.get("JoinStatus_2",""))
    js3 = sstatus(row.get("JoinStatus_3",""))
    join_stray = sstatus(row.get("JoinStray",""))

    allot1 = sval(row.get("Allot_1",""))
    allot2 = sval(row.get("Allot_2",""))
    allot3 = sval(row.get("Allot_3",""))
    straycol = sval(row.get("Stray",""))
    remarks = sval(row.get("Remarks",""))

    # ---------- SPECIAL RULE (highest priority)
    # If (joined in P1 or P2) AND (js3 = TC) AND (JoinStray = Y) -> full refund (regfeepaid)
    if (js1 == "Y" or js2 == "Y") and js3 == "TC" and join_stray == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Special: Joined P1/P2 + TC in P3 + Stray -> full refund"
        })

    # ---------- NEW RULE GROUP (priority after special)
    # RULE 2: No allotment anywhere AND joined via Spot (JoinStray=Y) AND actually joined in stray (js3='Y') -> full refund
    if allot1 == "" and allot2 == "" and allot3 == "" and join_stray == "Y" and js3 == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Rule2: No allotment & joined via Spot -> full refund"
        })

    # RULE 1: Paid again (forefit>0 & regfeepaid>0), no allotment in P3, joined via Spot -> refund second payment
    if existing_forefit > 0 and reg_paid > 0 and allot3 == "" and join_stray == "Y" and js3 == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Rule1: Paid again, no P3 allot, joined via spot -> refund 2nd payment"
        })

    # RULE 3: Paid again and GOT allotment in P3 and joined -> refund second payment
    if existing_forefit > 0 and reg_paid > 0 and allot3 != "" and js3 == "Y":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Rule3: Paid again, got P3 allot & joined -> refund 2nd payment"
        })

    # ---------- FALLBACK / STANDARD RULES
    # No allotment anywhere -> full refund (if no stray recorded)
    if allot1 == "" and allot2 == "" and allot3 == "" and straycol == "":
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "No allotment - full refund"
        })

    # Joined in phase 1 or 2 -> full refund unless TC (then forfeit)
    if js1 == "Y" or js2 == "Y":
        if "TC" in (js1, js2, js3):
            return pd.Series({
                "RefundAmount": 0.0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Joined then TC - no refund"
            })
        return pd.Series({
            "RefundAmount": reg_paid,
            "NewForefit": existing_forefit,
            "RefundCategory": "Joined in phase 1/2 - full refund"
        })

    # Phase 3 / Stray logic (if not joined in P1/P2)
    if js1 == "" and js2 == "":
        # joined in phase 3
        if js3 == "Y":
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": existing_forefit,
                "RefundCategory": "Joined in phase 3 - full refund"
            })
        # join via stray must be Y and js3 must be Y to refund (you said no refund if js3=N even if joinstray=Y)
        if join_stray == "Y" and js3 == "Y":
            return pd.Series({
                "RefundAmount": reg_paid,
                "NewForefit": existing_forefit,
                "RefundCategory": "Joined via stray - full refund"
            })
        if join_stray == "Y" and js3 != "Y":
            return pd.Series({
                "RefundAmount": 0.0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Stray allotted but NOT joined - no refund"
            })

    # Not joined / TC in phase 1 & 2 -> forfeit
    if js1 in ("N","TC") and js2 in ("N","TC"):
        if js3 == "N":
            return pd.Series({
                "RefundAmount": 0.0,
                "NewForefit": existing_forefit + reg_paid,
                "RefundCategory": "Not joined in 1/2/3 - forfeit"
            })
        return pd.Series({
            "RefundAmount": 0.0,
            "NewForefit": existing_forefit + reg_paid,
            "RefundCategory": "Not joined in 1/2 - forfeit"
        })

    # default: manual check
    return pd.Series({
        "RefundAmount": 0.0,
        "NewForefit": existing_forefit,
        "RefundCategory": "Check manually"
    })

# Apply computation
st.subheader("Computing refunds...")
calc = df.apply(compute_refund, axis=1)
df_out = pd.concat([df, calc], axis=1)

# Show results and summary
st.subheader("Result (first 200 rows)")
st.dataframe(df_out.head(200), use_container_width=True)

c1, c2, c3 = st.columns(3)
c1.metric("Total regfeepaid", f"{df_out['regfeepaid'].sum():,.2f}")
c2.metric("Total Refund (computed)", f"{df_out['RefundAmount'].sum():,.2f}")
c3.metric("Total NewForefit", f"{df_out['NewForefit'].sum():,.2f}")

# ----------------------
# Optional compare with manual file if provided
# ----------------------
if manual_file is not None:
    try:
        mfname = manual_file.name.lower()
        if mfname.endswith(".csv"):
            df_manual = pd.read_csv(manual_file, dtype=str)
        else:
            df_manual = pd.read_excel(manual_file, dtype=str)
    except Exception as e:
        st.error(f"Error reading manual file: {e}")
        df_manual = None

    if df_manual is not None:
        # normalize APPLNO and Refund column
        if "APPLNO" not in df_manual.columns:
            for alt in ["ApplNo","Applno","Appl_No"]:
                if alt in df_manual.columns:
                    df_manual.rename(columns={alt:"APPLNO"}, inplace=True)
                    break
        if "Refund" not in df_manual.columns:
            st.error("Manual file must include 'Refund' column to compare.")
        else:
            df_manual["APPLNO"] = df_manual["APPLNO"].astype(str).str.strip()
            df_manual["Refund"] = pd.to_numeric(df_manual["Refund"], errors="coerce").fillna(0)
            # merge
            merged = df_out.merge(df_manual[["APPLNO","Refund"]].rename(columns={"Refund":"ManualRefund"}),
                                  on="APPLNO", how="left")
            merged["ManualRefund"] = merged["ManualRefund"].fillna(0).astype(float)
            merged["ComputedRefund"] = merged["RefundAmount"].fillna(0).astype(float)
            merged["Diff"] = merged["ComputedRefund"] - merged["ManualRefund"]
            merged["Mismatch"] = merged["Diff"].abs() > 0.0001

            st.subheader("Comparison with manual file")
            st.write(f"Rows merged: {len(merged)}, mismatches: {int(merged['Mismatch'].sum())}")
            st.dataframe(merged[merged["Mismatch"]].head(200), use_container_width=True)
            # Save comparison files to disk for download
            out_all = merged.to_csv(index=False).encode("utf-8")
            out_mismatch = merged[merged["Mismatch"]].to_csv(index=False).encode("utf-8")

            st.download_button("Download full comparison CSV", data=out_all, file_name="refund_full_comparison_final.csv", mime="text/csv")
            st.download_button("Download mismatches CSV", data=out_mismatch, file_name="refund_mismatches_final.csv", mime="text/csv")

# ----------------------
# Export results: CSV (always) and colored Excel (if available)
# ----------------------
st.subheader("Export processed file")

csv_bytes = df_out.to_csv(index=False).encode("utf-8")
st.download_button("Download processed CSV", data=csv_bytes, file_name="refund_output.csv", mime="text/csv")

# Try Excel with colors
def make_colored_excel(df_to_write):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_to_write.to_excel(writer, index=False, sheet_name="Refund")
            workbook = writer.book
            worksheet = writer.sheets["Refund"]
            fmt_refund = workbook.add_format({"bg_color": "#c6efce"})
            fmt_forfeit = workbook.add_format({"bg_color": "#ffc7ce"})
            fmt_manual = workbook.add_format({"bg_color": "#ffeb9c"})
            # color rows
            for r in range(len(df_to_write)):
                cat = str(df_to_write.loc[r, "RefundCategory"]) if "RefundCategory" in df_to_write.columns else ""
                if "forfeit" in cat.lower():
                    worksheet.set_row(r+1, cell_format=fmt_forfeit)
                elif "refund" in cat.lower():
                    worksheet.set_row(r+1, cell_format=fmt_refund)
                elif "check manually" in cat.lower() or "check" in cat.lower():
                    worksheet.set_row(r+1, cell_format=fmt_manual)
        return output.getvalue()
    except Exception as e:
        return None

excel_bytes = make_colored_excel(df_out)
if excel_bytes:
    st.download_button("Download colored Excel (.xlsx)", data=excel_bytes, file_name="refund_output_colored.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.warning("Excel export not available in this environment (xlsxwriter missing). Use the CSV output.")

st.success("Processing complete. If any rows still mismatch your manual file, upload the manual file above and inspect the mismatches table; tell me any remaining special-case rules and I will integrate them.")
