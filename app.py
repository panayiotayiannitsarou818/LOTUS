# -*- coding: utf-8 -*-
import io, os, re, time, importlib.util, datetime as dt
from pathlib import Path

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Σχολικά Τμήματα — Wrapper", page_icon="🧩", layout="wide")

st.title("🧩 School Split — Thin Wrapper (Steps 1→7)")
st.caption("Δεν αλλάζει ΚΑΜΙΑ συνάρτηση στα modules. Απλό orchestration & export.")

ROOT = Path(__file__).parent

# -------- Helpers --------
def _load_module(name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(name, str(file_path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore
    return mod

def _check_required_files(required):
    missing = [str(p) for p in required if not p.exists()]
    return missing

def _read_file_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

def _timestamped(name: str, suffix: str) -> str:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"{name}_{ts}{suffix}"
    # sanitize for Excel sheet/file limits
    return re.sub(r"[^A-Za-z0-9_\-\.]+", "_", base)

# -------- File name consistency check --------
with st.expander("📦 Έλεγχος αρχείων / ονομάτων (consistency check)", expanded=True):
    needed = [
        ROOT / "export_step1_6_per_scenario.py",
        ROOT / "step1_immutable_ALLINONE.py",
        ROOT / "step_2_helpers_FIXED.py",
        ROOT / "step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED.py",
        ROOT / "step3_amivaia_filia_FIXED.py",
        ROOT / "step4_corrected.py",
        ROOT / "step5_enhanced.py",
        ROOT / "step6_compliant.py",
        ROOT / "step7_fixed_final.py",
    ]
    missing = _check_required_files(needed)
    if missing:
        st.error("❌ Λείπουν αρχεία (ονόματα/paths):\n" + "\n".join(f"- {m}" for m in missing))
    else:
        st.success("✅ Όλα τα απαραίτητα αρχεία βρέθηκαν με **συνεπή ονόματα**.")

# -------- Tabs --------
tab16, tab7 = st.tabs(["Βήματα 1→6", "Βήμα 7 (τελική επιλογή)"])

with tab16:
    st.subheader("Βήματα 1→6 — Παραγωγή σεναρίων per scenario")
    st.write("Χρησιμοποιείται **export_step1_6_per_scenario.build_step1_6_per_scenario**.")

    in_file = st.file_uploader("Ανέβασε το αρχικό Excel (input για Βήμα 1)", type=["xlsx"])
    colA, colB = st.columns(2)
    with colA:
        pick_step4 = st.selectbox("Κανόνας επιλογής στο Βήμα 4", ["best", "first", "strict"], index=0,
                                  help="Περνά ως `pick_step4` στο build_step1_6_per_scenario.")
    with colB:
        out_name = st.text_input("Όνομα αρχείου εξόδου", value=_timestamped("STEP1_6_PER_SCENARIO", ".xlsx"))

    run16 = st.button("▶️ Εκτέλεση Βήματα 1→6")
    if run16:
        if in_file is None:
            st.warning("Πρώτα ανέβασε ένα Excel.")
        elif missing:
            st.error("Δεν είναι δυνατή η εκτέλεση: λείπουν modules.")
        else:
            # Save upload
            input_path = ROOT / _timestamped("INPUT_STEP1", ".xlsx")
            with open(input_path, "wb") as f:
                f.write(in_file.getbuffer())

            # Import orchestrator
            m = _load_module("export_step1_6_per_scenario", ROOT / "export_step1_6_per_scenario.py")

            out_path = ROOT / out_name
            try:
                with st.spinner("Τρέχουν τα Βήματα 1→6..."):
                    m.build_step1_6_per_scenario(str(input_path), str(out_path), pick_step4=pick_step4)
                st.success("ΟΚ — ολοκληρώθηκε η παραγωγή των σεναρίων Βήμα 1→6.")
                st.download_button("⬇️ Κατέβασε το Excel (1→6)", data=_read_file_bytes(out_path),
                                   file_name=out_path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.exception(e)

with tab7:
    st.subheader("Βήμα 7 — Τελική κατάταξη & επιλογή")
    st.write("Χρησιμοποιείται το **step7_fixed_final.py** (χωρίς τροποποιήσεις).")

    step6_file = st.file_uploader("Ανέβασε το Excel από Βήμα 6 (per scenario)", type=["xlsx"], key="s6")
    col1, col2, col3 = st.columns(3)
    with col1:
        seed = st.number_input("Random seed για ισοβαθμίες", min_value=0, value=42, step=1)
    with col2:
        scores_name = st.text_input("Όνομα αρχείου Scores", value=_timestamped("STEP7_SCORES", ".xlsx"))
    with col3:
        final_name = st.text_input("Όνομα αρχείου Τελικού Σεναρίου", value=_timestamped("STEP7_FINAL_SCENARIO", ".xlsx"))

    run7 = st.button("🏁 Εκτέλεση Βήματος 7 (scoring & επιλογή)")
    if run7:
        if step6_file is None:
            st.warning("Πρώτα ανέβασε το Excel από το Βήμα 6.")
        elif missing:
            st.error("Δεν είναι δυνατή η εκτέλεση: λείπουν modules.")
        else:
            # Save upload
            s6_path = ROOT / _timestamped("INPUT_STEP6", ".xlsx")
            with open(s6_path, "wb") as f:
                f.write(step6_file.getbuffer())

            # Load modules
            s7 = _load_module("step7_fixed_final", ROOT / "step7_fixed_final.py")
            import pandas as pd

            # Read scenarios
            xls = pd.ExcelFile(s6_path)
            sheet_names = [s for s in xls.sheet_names if s != "Σύνοψη"]
            if not sheet_names:
                st.error("Δεν βρέθηκαν sheets σεναρίων (εκτός από 'Σύνοψη').")
            else:
                # Build a unified DF with one chosen scenario sheet (any; we only need columns)
                # We'll score by scanning scenario columns in the first sheet.
                df = pd.read_excel(s6_path, sheet_name=sheet_names[0])
                # Find all scenario columns ΒΗΜΑ6_ΣΕΝΑΡΙΟ_k
                scen_cols = [c for c in df.columns if re.match(r"^ΒΗΜΑ6_ΣΕΝΑΡΙΟ_\d+$", str(c))]
                if not scen_cols:
                    st.error("Δεν βρέθηκαν στήλες τύπου 'ΒΗΜΑ6_ΣΕΝΑΡΙΟ_N'.")
                else:
                    try:
                        # 1) Export scores
                        scores_out = ROOT / scores_name
                        s7.export_scores_excel(df.copy(), scen_cols, str(scores_out))

                        # 2) Pick best & build final workbook
                        pick = s7.pick_best_scenario(df.copy(), scen_cols, random_seed=int(seed))
                        best = pick.get("best")
                        if not best or "scenario_col" not in best:
                            st.error("Αποτυχία επιλογής σεναρίου.")
                        else:
                            winning_col = best["scenario_col"]
                            final_df = pd.read_excel(s6_path, sheet_name=sheet_names[0]).copy()

                            # Δημιουργία Excel: FINAL_SCENARIO + ανά τμήμα
                            final_out = ROOT / final_name
                            with pd.ExcelWriter(final_out, engine="xlsxwriter") as w:
                                final_df.to_excel(w, index=False, sheet_name="FINAL_SCENARIO")
                                # Sheets ανά τμήμα (Α1, Α2, ...), μόνο ονόματα
                                labels = sorted([str(v) for v in final_df[winning_col].dropna().unique() if re.match(r"^Α\d+$", str(v))],
                                                key=lambda x: int(re.search(r"\d+", x).group(0)))
                                for lab in labels:
                                    sub = final_df.loc[final_df[winning_col]==lab, ["ΟΝΟΜΑ", winning_col]].copy()
                                    sub = sub.rename(columns={winning_col: "ΤΜΗΜΑ"})
                                    sub.to_excel(w, index=False, sheet_name=str(lab))

                            st.success(f"Νικητής: στήλη {winning_col}")
                            st.download_button("⬇️ Κατέβασε Scores (Βήμα 7)",
                                               data=_read_file_bytes(scores_out),
                                               file_name=scores_out.name,
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            st.download_button("⬇️ Κατέβασε Τελικό Σενάριο",
                                               data=_read_file_bytes(final_out),
                                               file_name=final_out.name,
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.exception(e)

st.divider()
st.caption("Wrapper μόνο — δεν αγγίζει business logic. Αν αλλάξουν τα modules, ο wrapper απλά τα ξαναφορτώνει.")
