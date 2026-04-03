"""
NHS Hospital Admissions Data Cleaning Script v2
COMP4037 Research Methods - Coursework 2
Compatible with three formats: all years from 1998-99 to 2023-24

"""

import pandas as pd
import os
import re

# ============================================================
# The folder path where all Excel files are stored
# ============================================================
DATA_FOLDER = r".\hospital_data"

# ============================================================
# File configuration table
# ============================================================
FILE_CONFIG = {
    "1998-99": ("hosp-epis-stat-admi-prim-diag-sum-98-99-tab.xlsx", "Diagnosis199899"),
    "1999-00": ("hosp-epis-stat-admi-prim-diag-sum-99-00-tab.xlsx", "Diagnosis9900"),
    "2000-01": ("hosp-epis-stat-admi-prim-diag-sum-00-01-tab.xlsx", "Diagnosis0001"),
    "2001-02": ("hosp-epis-stat-admi-prim-diag-sum-01-02-tab.xlsx", "Diagnosis0102"),
    "2002-03": ("hosp-epis-stat-admi-prim-diag-sum-02-03-tab.xlsx", "Primary Diagnosis - summary"),
    "2003-04": ("hosp-epis-stat-admi-prim-diag-sum-03-04-tab.xls.xlsx", "Primary Diagnosis - summary"),
    "2004-05": ("hosp-epis-stat-admi-prim-diag-sum-04-05-tab.xls.xlsx", "Primary Diagnosis Summary"),
    "2005-06": ("hosp-epis-stat-admi-prim-diag-sum-05-06-tab.xlsx", "Primary diagnosis - summary"),
    "2006-07": ("hosp-epis-stat-admi-prim-diag-sum-06-07-tab.xlsx", "Primary diagnosis - summary"),
    "2007-08": ("hosp-epis-stat-admi-prim-diag-07-08-tab.xlsx", "Primary diagnosis - summary"),
    "2008-09": ("hosp-epis-stat-admi-prim-diag-sum-08-09-tab.xlsx", "Primary diagnosis - summary"),
    "2009-10": ("hosp-epis-stat-admi-prim-diag-sum-09-10-tab.xlsx", "Primary diagnosis - summary"),
    "2010-11": ("hosp-epis-stat-admi-prim-2010-11-diag-sum.xlsx", "Primary diagnosis - summary"),
    "2011-12": ("hosp-epis-stat-admi-diag-2011-12-prim-diag-sum.xlsx", "Primary diagnosis - summary"),
    "2012-13": ("hosp-epis-stat-admi-diag-2012-13-tab.xlsx", "Primary diagnosis - summary"),
    "2013-14": ("hosp-epis-stat-admi-diag-2013-14-tab.xlsx", "Primary Diagnosis Summary"),
    "2014-15": ("hosp-epis-stat-admi-diag-2014-15-tab.xlsx", "Primary Diagnosis Summary"),
    "2015-16": ("hosp-epis-stat-admi-diag-2015-16-tab.xlsx", "Primary Diagnosis Summary"),
    "2016-17": ("hosp-epis-stat-admi-diag-2016-17-tab.xlsx", "Primary Diagnosis Summary"),
    "2017-18": ("hosp-epis-stat-admi-diag-2017-18-tab.xlsx", "Primary Diagnosis Summary"),
    "2018-19": ("hosp-epis-stat-admi-diag-2018-19-tab.xlsx", "Primary Diagnosis Summary"),
    "2019-20": ("hosp-epis-stat-admi-diag-2019-20-tab supp.xlsx", "Primary Diagnosis Summary"),
    "2020-21": ("hosp-epis-stat-admi-diag-2020-21-tab.xlsx", "Primary Diagnosis Summary"),
    "2021-22": ("hosp-epis-stat-admi-diag-2021-22-tab.xlsx", "Primary Diagnosis Summary"),
    "2022-23": ("hosp-epis-stat-admi-diag-2022-23-tab_V2.xlsx", "Primary Diagnosis Summary"),
    "2023-24": ("hosp-epis-stat-admi-diag-2023-24-tab.xlsx", "Primary Diagnosis Summary"),
}


# ============================================================
# ICD chapter mapping
# ============================================================
ICD_CHAPTERS = {
    "A": "Ch.I  Infectious & parasitic",
    "B": "Ch.I  Infectious & parasitic",
    "C": "Ch.II  Neoplasms",
    "D": "Ch.II  Neoplasms",
    "E": "Ch.IV  Endocrine & metabolic",
    "F": "Ch.V  Mental & behavioural",
    "G": "Ch.VI  Nervous system",
    "H": "Ch.VII/VIII  Eye & Ear",
    "I": "Ch.IX  Circulatory system",
    "J": "Ch.X  Respiratory system",
    "K": "Ch.XI  Digestive system",
    "L": "Ch.XII  Skin",
    "M": "Ch.XIII  Musculoskeletal",
    "N": "Ch.XIV  Genitourinary",
    "O": "Ch.XV  Pregnancy & childbirth",
    "P": "Ch.XVI  Perinatal",
    "Q": "Ch.XVII  Congenital",
    "R": "Ch.XVIII  Symptoms & signs",
    "S": "Ch.XIX  Injury & poisoning",
    "T": "Ch.XIX  Injury & poisoning",
    "V": "Ch.XX  External causes",
    "W": "Ch.XX  External causes",
    "X": "Ch.XX  External causes",
    "Y": "Ch.XX  External causes",
    "Z": "Ch.XXI  Health status factors",
}


# ============================================================
# built-in function
# ============================================================

def find_header_row(df_raw):
    """Find the title line containing the keyword 'admission' and within the first 30 lines"""
    for i, row in df_raw.iterrows():
        row_str = " ".join(str(v) for v in row if pd.notna(v)).lower()
        if "admission" in row_str and i < 30:
            return i
    return None


def detect_format(df_raw, header_row):
    """
    Judge the data format
        'merged' → Diagnostic code and description merged in the first column (1998-99-2012-13)
        'split' → Diagnostic codes and descriptions are in two columns (2013-14-2023-24)
    """
    for i in range(header_row + 1, min(header_row + 6, len(df_raw))):
        cell = str(df_raw.iloc[i, 0]).strip()
        if cell.lower() == "total" or cell == "nan":
            continue
        # Pure code: A00-A09 or A00 (short length)
        if re.match(r'^[A-Z]\d{2}', cell) and len(cell) <= 12:
            return 'split'
        # mergeformat: A00-A09 Intestinal infectious...
        if re.match(r'^[A-Z‡†]', cell) and len(cell) > 12:
            return 'merged'
    return 'split'


def map_columns(df):
    """
    Scan all column names and return the field mapping dictionary.
    Unified output field：admissions, fce, emergency, waiting_list,
                  planned, male, female, mean_age, mean_los, median_los
    """
    col_map = {}
    for c in df.columns:
        # Clean column names: Remove line breaks, first and last Spaces, and convert to lowercase
        cl = str(c).lower().replace('\n', ' ').replace('\r', ' ').strip()
        cl = re.sub(r'\s+', ' ', cl)

        # admissions（Total Hospitalization）：Must be ruled out with emergency/waiting/planned/other/elective column
        if 'admissions' not in col_map:
            if re.search(r'finished admission episode', cl):
                col_map['admissions'] = c
            elif re.search(r'^admissions$', cl):
                col_map['admissions'] = c

        # fce (Completed medical sessions)
        if 'fce' not in col_map:
            if re.search(r'finished consultant episode', cl):
                col_map['fce'] = c

        # emergency (Emergency Hospitalization) : Exclude columns containing zero/bed/.1 (duplicate columns)
        if 'emergency' not in col_map:
            if (re.search(r'\bemergency\b', cl)
                    and 'zero' not in cl
                    and 'bed' not in cl
                    and not cl.endswith('.1')):
                col_map['emergency'] = c

        # waiting_list
        if 'waiting_list' not in col_map:
            if re.search(r'waiting list', cl):
                col_map['waiting_list'] = c

        # planned
        if 'planned' not in col_map:
            if re.search(r'\bplanned\b', cl) and 'un' not in cl:
                col_map['planned'] = c

        # male / female
        if 'male' not in col_map:
            if re.match(r'^male', cl):
                col_map['male'] = c
        if 'female' not in col_map:
            if re.match(r'^female', cl):
                col_map['female'] = c

        # mean_age
        if 'mean_age' not in col_map:
            if re.search(r'mean age', cl):
                col_map['mean_age'] = c

        # mean_los / median_los
        if 'mean_los' not in col_map:
            if re.search(r'mean length', cl):
                col_map['mean_los'] = c
        if 'median_los' not in col_map:
            if re.search(r'median length', cl):
                col_map['median_los'] = c

    return col_map


def split_code_description(cell):
    """'A00-A09 Intestinal infectious diseases' split into('A00-A09', 'Intestinal...')"""
    cell = str(cell).strip()
    # Try: The code part (letter + number + hyphen + comma) + two or more Spaces + description
    m = re.match(r'^([‡†\s]*[A-Z][\d\-,A-Z\s]+?)\s{2,}(.+)$', cell)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # Backup: Code + one space + description (the description should be at least 4 characters)
    m2 = re.match(r'^([A-Z‡†][\w\-,]+)\s+(.{4,})$', cell)
    if m2:
        return m2.group(1).strip(), m2.group(2).strip()
    return cell, ""


def read_one_year(filepath, sheet_name, year_label):
    """Read the single-year Summary sheet and return the standardized DataFrame"""
    print(f"  {year_label} ...", end="  ")

    # Read the original data (all as strings to avoid type conversion issues)
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name,
                           header=None, dtype=str)

    # Find the title line
    header_row = find_header_row(df_raw)
    if header_row is None:
        print("❌ The title line cannot be found")
        return None

    # Judgment format
    fmt = detect_format(df_raw, header_row)

    # Re-read with the title line
    df = pd.read_excel(filepath, sheet_name=sheet_name,
                       header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # ── Handle diagnostic codes and descriptions──────────────────────────
    if fmt == 'merged':
        first_col = df.columns[0]
        parsed = df[first_col].apply(split_code_description)
        df.insert(0, '__desc__', parsed.apply(lambda x: x[1]))
        df.insert(0, '__code__', parsed.apply(lambda x: x[0]))
        df = df.drop(columns=[first_col])
        code_col, desc_col = '__code__', '__desc__'
    else:
        code_col = df.columns[0]
        desc_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    # ───Map the remaining columns────────────────────────────────
    col_map = map_columns(df)

    # ── Assembly output DataFrame ──────────────────────────────
    out = {
        'diag_code':   df[code_col],
        'description': df[desc_col],
    }
    for field in ['fce', 'admissions', 'emergency', 'waiting_list',
                  'planned', 'male', 'female', 'mean_age', 'mean_los', 'median_los']:
        out[field] = df[col_map[field]] if field in col_map else pd.NA

    result = pd.DataFrame(out)

    # ── clean line ──────────────────────────────────────────
    result['diag_code']   = result['diag_code'].astype(str).str.strip()
    result['description'] = result['description'].astype(str).str.strip()

    # Only diagnostic lines starting with a capital letter are retained (excluding Total, blank lines, and comment lines)
    result = result[result['diag_code'].str.match(r'^[A-Z‡†]', na=False)].copy()

    # Remove the annotation symbols such as  ‡ †
    result['diag_code'] = (result['diag_code']
                           .str.replace(r'^[‡†\s]+', '', regex=True)
                           .str.strip())

    # Filter again (blank lines may appear after removing the comment symbols)
    result = result[result['diag_code'].str.match(r'^[A-Z]', na=False)].copy()

    # ── Digital column conversion ──────────────────────────────────────
    num_cols = ['fce', 'admissions', 'emergency', 'waiting_list',
                'planned', 'male', 'female', 'mean_age', 'mean_los', 'median_los']
    for col in num_cols:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors='coerce')

    # ── Derived field ────────────────────────────────────────
    if 'emergency' in result.columns and 'admissions' in result.columns:
        result['emergency_ratio'] = (result['emergency'] / result['admissions']).round(4)
    if 'male' in result.columns and 'admissions' in result.columns:
        result['male_ratio'] = (result['male'] / result['admissions']).round(4)

    # ── metadata──────────────────────────────────────────
    result['year']       = year_label
    result['year_start'] = int(year_label[:4])
    result['icd_chapter']= result['diag_code'].str[0].map(ICD_CHAPTERS).fillna('Other')
    result['format']     = fmt

    # Count the number of valid fields (for easy debugging)
    filled = sum(result[c].notna().any() for c in num_cols if c in result.columns)
    print(f"✅  {len(result):>3} item  |  format:{fmt}  |  numeric field:{filled}/{len(num_cols)}")
    return result


# ============================================================
# main program
# ============================================================
def main():
    print("=" * 65)
    print("NHS Hospital Admissions Data Cleaning v2  (1998–2024)")
    print("=" * 65)

    all_dfs = []

    for year_label, (filename, sheet_name) in sorted(FILE_CONFIG.items()):
        filepath = os.path.join(DATA_FOLDER, filename)
        if not os.path.exists(filepath):
            print(f"  ⚠️  The file does not exist. Skip: {filename}")
            continue
        try:
            df = read_one_year(filepath, sheet_name, year_label)
            if df is not None:
                all_dfs.append(df)
        except Exception as e:
            print(f"  ❌ {year_label} defeat: {e}")

    if not all_dfs:
        print("❌ No files were successfully read. Please check the path")
        return

    # ── combine ──────────────────────────────────────────
    print(f"\n combine {len(all_dfs)} years ...")
    combined = pd.concat(all_dfs, ignore_index=True)
    combined = combined.sort_values(['year_start', 'diag_code']).reset_index(drop=True)
    print(f"✅Total {len(combined):,} rows  × {len(combined.columns)} columns")
    print(f"   Years：{combined['year'].min()} → {combined['year'].max()}")

    # ── Output 1: Complete long table ─────────────────────────────────
    p1 = os.path.join(DATA_FOLDER, "nhs_combined_all_years.csv")
    combined.to_csv(p1, index=False, encoding="utf-8-sig")
    print(f"\n📄 Output 1 Complete data:         {p1}")

    # ── Aggregate by chapter × year ─────────────────────────────────
    agg_cols = {c: 'sum' for c in
                ['fce','admissions','emergency','waiting_list','planned','male','female']
                if c in combined.columns}
    chapter_yr = (combined
                  .groupby(['year', 'year_start', 'icd_chapter'])
                  .agg(agg_cols)
                  .reset_index()
                  .sort_values(['year_start', 'icd_chapter']))
    if 'emergency' in chapter_yr and 'admissions' in chapter_yr:
        chapter_yr['emergency_ratio'] = (chapter_yr['emergency'] /
                                         chapter_yr['admissions']).round(4)

    p2 = os.path.join(DATA_FOLDER, "nhs_by_chapter_year.csv")
    chapter_yr.to_csv(p2, index=False, encoding="utf-8-sig")
    print(f"📄 Output 2 chapters × year aggregation:     {p2}")

    # Year sequence (used for sorting heat map columns)
    year_order = (chapter_yr.drop_duplicates('year')
                             .sort_values('year_start')['year'].tolist())

    # ── Output 3: Wide heat map (Number of emergency patients)──────────────────
    if 'emergency' in chapter_yr.columns:
        piv = chapter_yr.pivot_table(
            index='icd_chapter', columns='year',
            values='emergency', aggfunc='sum')
        piv = piv[[y for y in year_order if y in piv.columns]]
        p3 = os.path.join(DATA_FOLDER, "nhs_heatmap_emergency.csv")
        piv.to_csv(p3, encoding="utf-8-sig")
        print(f"📄 Output 3 heat maps - Emergency:       {p3}")

    # ── Output 4: Wide heat map (total number of inpatients)───────────────
    if 'admissions' in chapter_yr.columns:
        piv2 = chapter_yr.pivot_table(
            index='icd_chapter', columns='year',
            values='admissions', aggfunc='sum')
        piv2 = piv2[[y for y in year_order if y in piv2.columns]]
        p4 = os.path.join(DATA_FOLDER, "nhs_heatmap_admissions.csv")
        piv2.to_csv(p4, encoding="utf-8-sig")
        print(f"📄 Output 4 heat maps - Total hospitalization:     {p4}")

    # ── Output 5: Comparison before and after COVID───────────────────────────
    covid_years = ['2018-19', '2019-20', '2020-21', '2021-22', '2022-23']
    covid = chapter_yr[chapter_yr['year'].isin(covid_years)]
    p5 = os.path.join(DATA_FOLDER, "nhs_covid_impact.csv")
    covid.to_csv(p5, index=False, encoding="utf-8-sig")
    print(f"📄 Output 5 comparisons before and after COVID:     {p5}")

    # ──  Console summary─────────────────────────────────────
    print("\n" + "=" * 65)
    print("Data summary")
    print("=" * 65)
    print(f"Year：{combined['year'].nunique()}  |  "
          f"Diagnostic category:{combined['icd_chapter'].nunique()}  |  "
          f"The years with emergency data：{combined[combined['emergency'].notna()]['year'].nunique()}")

    if 'admissions' in combined.columns:
        print("\nTotal number of inpatients in each year:")
        yr_tot = (combined.groupby(['year_start','year'])['admissions']
                          .sum().reset_index().sort_values('year_start'))
        for _, row in yr_tot.iterrows():
            bar = '█' * int(row['admissions'] / 600_000)
            print(f"  {row['year']}  {bar}  ({row['admissions']/1e6:.1f}M)")

    if 'emergency' in chapter_yr.columns:
        pre  = chapter_yr[chapter_yr['year']=='2019-20'].set_index('icd_chapter')['emergency']
        post = chapter_yr[chapter_yr['year']=='2020-21'].set_index('icd_chapter')['emergency']
        change = ((post - pre) / pre * 100).dropna().sort_values()
        if not change.empty:
            print("\n2020-21 vs 2019-20 Emergency department changes (Impact of COVID lockdowns：")
            for cat, pct in change.items():
                arrow = "▼" if pct < 0 else "▲"
                print(f"  {arrow} {pct:+5.1f}%  {cat}")

    print("\n✅ All done! Please check the generated CSV file.CSV 文件")


if __name__ == "__main__":
    main()
