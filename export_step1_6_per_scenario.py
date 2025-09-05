# -*- coding: utf-8 -*-
"""
export_step1_6_per_scenario.py — ΔΙΟΡΘΩΜΕΝΟΣ exporter (1→6)

Εκθέτει τη συνάρτηση:
    build_step1_6_per_scenario(input_excel, output_excel, pick_step4="best")

Τρέχει ΟΛΟΚΛΗΡΗ τη ροή: Βήματα 1→6
"""

from typing import Optional, List, Tuple
import importlib.util, sys, re, numpy as np, pandas as pd
from pathlib import Path

CORE_COLUMNS = [
    "ΟΝΟΜΑ","ΦΥΛΟ","ΖΩΗΡΟΣ","ΙΔΙΑΙΤΕΡΟΤΗΤΑ","ΠΑΙΔΙ_ΕΚΠΑΙΔΕΥΤΙΚΟΥ",
    "ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ","ΦΙΛΟΙ","ΣΥΓΚΡΟΥΣΗ"
]

def _import(modname: str, path: Path):
    spec = importlib.util.spec_from_file_location(modname, str(path))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[modname] = mod
    spec.loader.exec_module(mod)
    return mod

def _sid(col: str) -> int:
    m = re.search(r"ΣΕΝΑΡΙΟ[_\s]*(\d+)", str(col))
    return int(m.group(1)) if m else 1

def _dedup(df: pd.DataFrame) -> pd.DataFrame:
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
    return df

def build_step1_6_per_scenario(input_excel: str, output_excel: str, pick_step4: str = "best") -> None:
    root = Path(__file__).parent
    
    # Import όλων των modules
    m_step1 = _import("step1_immutable_ALLINONE", root / "step1_immutable_ALLINONE.py")
    m_help2 = _import("step_2_helpers_FIXED", root / "step_2_helpers_FIXED.py")
    m_step2 = _import("step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED", root / "step_2_zoiroi_idiaterotites_FIXED_v3_PATCHED.py")
    m_h3    = _import("step3_amivaia_filia_FIXED", root / "step3_amivaia_filia_FIXED.py")
    m_step4 = _import("step4_corrected", root / "step4_corrected.py")
    m_step5 = _import("step5_enhanced", root / "step5_enhanced.py")
    m_step6 = _import("step6_compliant", root / "step6_compliant.py")

    # Συμβατότητα υπογραφής στο Step4
    if hasattr(m_step4, "count_groups_by_category_per_class_strict"):
        _orig = m_step4.count_groups_by_category_per_class_strict
        def _count_wrapper(df, assigned_column, classes, step1_results=None, detected_pairs=None):
            return _orig(df, assigned_column, classes, step1_results, detected_pairs)
        m_step4.count_groups_by_category_per_class_strict = _count_wrapper

    xls = pd.ExcelFile(input_excel)
    df0 = xls.parse(xls.sheet_names[0])

    # STEP 1
    df1, _ = m_step1.create_immutable_step1(df0, num_classes=None)

    # Κενά -> NaN
    for c in [c for c in df1.columns if str(c).startswith("ΒΗΜΑ1_ΣΕΝΑΡΙΟ_")]:
        mask = df1[c].astype(str).str.strip() == ""
        df1.loc[mask, c] = np.nan

    step1_cols = sorted(
        [c for c in df1.columns if str(c).startswith("ΒΗΜΑ1_ΣΕΝΑΡΙΟ_")],
        key=_sid
    )

    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as w:
        for s1col in step1_cols:
            sid = _sid(s1col)

            # STEP 2
            options2 = m_step2.step2_apply_FIXED_v3(df1.copy(), step1_col_name=s1col, seed=42, max_results=5)
            if options2:
                df2 = options2[0][1]
                s2col = f"ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{sid}"
                if s2col not in df2.columns:
                    cands = [c for c in df2.columns if str(c).startswith("ΒΗΜΑ2_")]
                    s2col = cands[0] if cands else s2col
                    if s2col not in df2.columns:
                        df2[s2col] = ""
            else:
                df2 = df1.copy(); s2col = f"ΒΗΜΑ2_ΣΕΝΑΡΙΟ_{sid}"; df2[s2col] = ""

            base = df1.copy()
            base = base.merge(df2[["ΟΝΟΜΑ", s2col]], on="ΟΝΟΜΑ", how="left")

            # Βάλε τη ΒΗΜΑ2 δίπλα στη ΒΗΜΑ1
            cols = base.columns.tolist()
            if s2col in cols: cols.remove(s2col)
            idx = cols.index(s1col) + 1 if s1col in cols else len(cols)
            cols = cols[:idx] + [s2col] + cols[idx:]
            base = base[cols]

            # STEP 3
            df3, _ = m_h3.apply_step3_on_sheet(base.copy(), scenario_col=s2col, num_classes=None)
            s3col = f"ΒΗΜΑ3_ΣΕΝΑΡΙΟ_{sid}"
            cands3 = [c for c in df3.columns if str(c).startswith("ΒΗΜΑ3_")]
            if cands3 and s3col not in cands3:
                df3 = df3.rename(columns={cands3[0]: s3col})
            elif s3col not in df3.columns:
                df3[s3col] = ""

            # Βάλε τη ΒΗΜΑ3 δίπλα στη ΒΗΜΑ2
            cols3 = df3.columns.tolist()
            if s3col in cols3: cols3.remove(s3col)
            idx2 = cols3.index(s2col) + 1 if s2col in cols3 else len(cols3)
            cols3 = cols3[:idx2] + [s3col] + cols3[idx2:]
            df3 = df3[cols3]

            # Προετοιμασία ΦΙΛΟΙ για Step4
            if "ΦΙΛΟΙ" in df3.columns:
                try:
                    df3["ΦΙΛΟΙ"] = df3["ΦΙΛΟΙ"].apply(m_help2.parse_friends_cell)
                except Exception:
                    pass

            # STEP 4
            res4 = m_step4.apply_step4_with_enhanced_strategy(
                df3.copy(), assigned_column=s3col, num_classes=None, max_results=5
            )
            s4final = f"ΒΗΜΑ4_ΣΕΝΑΡΙΟ_{sid}"
            if res4:
                df4_mat = m_step4.export_step4_scenarios(df3.copy(), res4, assigned_column=s3col)
                if str(pick_step4).lower() == "best":
                    penalties = [p for (_, p) in res4]
                    best_idx = int(min(range(len(penalties)), key=lambda i: penalties[i]))
                    src = f"ΒΗΜΑ4_ΣΕΝΑΡΙΟ_{best_idx+1}"
                else:
                    try:
                        idx_pick = max(1, min(int(pick_step4), len(res4)))
                    except Exception:
                        idx_pick = 1
                    src = f"ΒΗΜΑ4_ΣΕΝΑΡΙΟ_{idx_pick}"
                cands4 = [c for c in df4_mat.columns if str(c).startswith("ΒΗΜΑ4_")]
                if src in df4_mat.columns:
                    df4 = df4_mat.rename(columns={src: s4final})
                elif cands4:
                    df4 = df4_mat.rename(columns={cands4[0]: s4final})
                else:
                    df4 = df3.copy(); df4[s4final] = ""
            else:
                df4 = df3.copy(); df4[s4final] = ""

            # Βάλε τη ΒΗΜΑ4 δίπλα στη ΒΗΜΑ3
            cols4 = df4.columns.tolist()
            if s4final in cols4: cols4.remove(s4final)
            idx3 = cols4.index(s3col) + 1 if s3col in cols4 else len(cols4)
            cols4 = cols4[:idx3] + [s4final] + cols4[idx3:]
            df4 = df4[cols4]
            df4 = _dedup(df4)

            # STEP 5
            df5, _pen5 = m_step5.step5_place_remaining_students(df4.copy(), scenario_col=s4final, num_classes=None)
            s5col = f"ΒΗΜΑ5_ΣΕΝΑΡΙΟ_{sid}"
            df5[s5col] = df5[s4final]
            cols5 = df5.columns.tolist()
            if s5col in cols5: cols5.remove(s5col)
            idx4 = cols5.index(s4final) + 1 if s4final in cols5 else len(cols5)
            cols5 = cols5[:idx4] + [s5col] + cols5[idx4:]
            df5 = df5[cols5]

            # STEP 6 - ΠΡΟΣΘΗΚΗ
            # Προετοιμασία δεδομένων για Step 6
            df5_prep = df5.copy()
            if "Α/Α" not in df5_prep.columns:
                df5_prep["Α/Α"] = range(1, len(df5_prep) + 1)
            if "ΤΜΗΜΑ_ΒΗΜΑ1" not in df5_prep.columns: 
                df5_prep["ΤΜΗΜΑ_ΒΗΜΑ1"] = df5_prep[s1col]
            if "ΤΜΗΜΑ_ΒΗΜΑ2" not in df5_prep.columns: 
                df5_prep["ΤΜΗΜΑ_ΒΗΜΑ2"] = df5_prep[s2col]
            if "GROUP_ID" not in df5_prep.columns: 
                df5_prep["GROUP_ID"] = np.nan
            if "ΒΗΜΑ_ΤΟΠΟΘΕΤΗΣΗΣ" not in df5_prep.columns:
                df5_prep["ΒΗΜΑ_ΤΟΠΟΘΕΤΗΣΗΣ"] = [
                    4 if str(l).strip() != "" else (5 if str(m).strip() != "" else np.nan) 
                    for l, m in zip(df5_prep[s4final], df5_prep[s5col])
                ]

            # Εκτέλεση Step 6
            try:
                step6_result = m_step6.apply_step6(
                    df5_prep.copy(),
                    class_col=s5col, 
                    id_col="Α/Α",
                    gender_col="ΦΥΛΟ", 
                    lang_col="ΚΑΛΗ_ΓΝΩΣΗ_ΕΛΛΗΝΙΚΩΝ",
                    step_col="ΒΗΜΑ_ΤΟΠΟΘΕΤΗΣΗΣ", 
                    group_col="GROUP_ID",
                    max_iter=5
                )
                df6 = step6_result["df"]
                
                s6col = f"ΒΗΜΑ6_ΣΕΝΑΡΙΟ_{sid}"
                # Χρήση του τελικού αποτελέσματος από Step 6
                if "ΒΗΜΑ6_ΤΜΗΜΑ" in df6.columns:
                    df6[s6col] = df6["ΒΗΜΑ6_ΤΜΗΜΑ"]
                elif f"ΒΗΜΑ6_ΣΕΝΑΡΙΟ_{sid}" in df6.columns:
                    pass  # Ήδη υπάρχει
                else:
                    df6[s6col] = df6[s5col]  # Fallback
                
                # Βάλε τη ΒΗΜΑ6 δίπλα στη ΒΗΜΑ5
                cols6 = df6.columns.tolist()
                if s6col in cols6: cols6.remove(s6col)
                idx5 = cols6.index(s5col) + 1 if s5col in cols6 else len(cols6)
                cols6 = cols6[:idx5] + [s6col] + cols6[idx5:]
                df6 = df6[cols6]
                
            except Exception as e:
                print(f"Σφάλμα στο Step 6 για σενάριο {sid}: {e}")
                df6 = df5.copy()
                s6col = f"ΒΗΜΑ6_ΣΕΝΑΡΙΟ_{sid}"
                df6[s6col] = df6[s5col]  # Fallback: ΒΗΜΑ6 = ΒΗΜΑ5

            # Κράτα CORE στήλες + όλα τα βήματα
            keep = [c for c in CORE_COLUMNS if c in df6.columns] + [s1col, s2col, s3col, s4final, s5col, s6col]
            out_df = _dedup(df6[keep].copy())

            sheet_name = f"ΣΕΝΑΡΙΟ_{sid}"
            out_df.to_excel(w, sheet_name=sheet_name[:31], index=False)

# Aliases για συμβατότητα
build_step1_4_per_scenario = build_step1_6_per_scenario
build_step1_5_per_scenario = build_step1_6_per_scenario
