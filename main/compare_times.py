import sys
import subprocess

# Automatically install missing external dependencies
def install_and_import(package, import_name=None):
    """Automatically installs missing packages via pip if not found."""
    if import_name is None:
        import_name = package
    try:
        __import__(import_name)
    except ImportError:
        print(f" Package '{package}' missing. Installing automatically using pip...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_and_import("pandas")
install_and_import("openpyxl")

import re
from pathlib import Path
from tkinter import Tk, filedialog
import pandas as pd

def select_files_via_gui() -> list[Path]:
    """Opens a native file selection window allowing multi-file select."""
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    print("="*60)
    print("\n Välj filer att jämföra. Programmet kommer identifiera gruppens tider och sammanfatta dagens tider för varje person.")
    print(" Denna databasen jämförs sedan med stämplade tider från Personec.")
    print(" För bästa upplevelse bör fil eller Excel-Sheet innehålla medarbetarens namn eller anställningsnummer.")
    print(" Kan inte medarbetare identifieras kommer användaren uppmanas mata in uppgiften.\n")
    print("="*60, "\n")
    
    file_paths = filedialog.askopenfilenames(
        title="Select the files to process",
        filetypes=[("All Files", "*.*")]
    )
    
    root.quit()
    root.destroy()
    return [Path(p) for p in file_paths]

def clear_temp_data(temp_dir: Path = Path("./temp_data")):
    if not temp_dir.exists():
        return
    for file_path in temp_dir.iterdir():
        if file_path.is_file():
            file_path.unlink()


def skip_to_header(file_path: Path) -> pd.DataFrame | None:
    target_key = "Datum"
    skip_amount = None

    try:
        with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
            for i, line in enumerate(f):
                if target_key in line:
                    skip_amount = i
                    break
        if skip_amount is None:
            return None

        df = pd.read_csv(file_path, skiprows=skip_amount, sep=None, engine="python", skipinitialspace=True, dtype=str)

        df.columns = [str(col).replace('\xa0', ' ').strip() for col in df.columns]
        df = df.map(lambda x: str(x).replace('\xa0', ' ').strip() if pd.notna(x) else None)
        df = df.replace({'': None, 'nan': None, 'None': None})
        is_unnamed = df.columns.str.startswith("Unnamed:") | (df.columns == "")
        df = df.loc[:, ~(is_unnamed & df.isna().all())]

        if "Medarbetare" in df.columns:
            cols = list(df.columns)
            idx = cols.index("Medarbetare") + 1
            if idx < len(cols) and cols[idx].startswith("Unnamed"):
                df = df.rename(columns={cols[idx]: "Namn"})

        return df.dropna(how="all").reset_index(drop=True)

    except Exception as e:
        print(f" [ERROR] Could not read file {file_path.name}: {e}")
        return None
    
    
def csv_process(file_path: Path, output_dir: Path):
    df = skip_to_header(file_path)   
    if df is None or df.empty:
        print(f" [DEBUG] Skipped {file_path.name} (no valid data found).")
        return

    base_name = file_path.stem  
    output_filename = f"{base_name}_eam.csv" if "AO-aktivitet" in df.columns and "eam" not in base_name.lower() else f"{base_name}.csv"
    df.to_csv(output_dir / output_filename, index=False, encoding="utf-8-sig")


def excel_to_csv(file_path: Path, output_dir: Path):
    base_name = file_path.stem  
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    except ImportError:
        print("\n [ERROR] 'openpyxl' is missing! Run: pip install openpyxl in your terminal.")
        return
    except Exception as e:
        print(f"\n [ERROR] Could not read {file_path.name}: {e}")
        return

    for sheet_name, df in excel_data.items():
        mask = df.isin(["Datum"]).any(axis=1)
        if mask.any():
            skip_idx = mask.idxmax()            
            df.columns = df.iloc[skip_idx]
            df = df.iloc[skip_idx + 1 :].reset_index(drop=True)

        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        if len(df.columns) > 1 and df.columns[0] == "Medarbetare" and pd.isna(df.columns[1]):    
            df = df.rename(columns={df.columns[1]: "Namn"})
        
        empty_header = df.columns.isna() | (df.columns.astype(str).str.strip() == "")
        empty_values = df.isna().all(axis=0)
        df = df.loc[:, ~(empty_header & empty_values)]
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
            
        output_filename = f"{base_name}_{sheet_name}_eam.csv" if "AO-aktivitet" in df.columns and "eam" not in base_name.lower() else f"{base_name}_{sheet_name}.csv"
        df.to_csv(output_dir / output_filename, index=False, encoding="utf-8")
        
        
def process_files(files: list[Path]):
    output_dir = Path("./temp_data")
    output_dir.mkdir(parents=True, exist_ok=True)

    for file in files:
        ext = file.suffix.lower()
        if ext in [".xlsx", ".xls"]:
            excel_to_csv(file, output_dir)
        elif ext == ".csv":
            csv_process(file, output_dir)
        else:
            print(f"\n[Skipped Unsupported File]: {file.name}")


def parse_job_times(file_path: Path) -> pd.DataFrame:
    df = skip_to_header(file_path)
    if df is None or df.empty:
        return pd.DataFrame()

    df["Medarbetare"] = df["Medarbetare"].ffill()
    df["Namn"] = df["Namn"].ffill()
    df = df.dropna(subset=["Datum", "Arbetsorder"]).copy()
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce").dt.strftime("%Y-%m-%d")
    df = df.dropna(subset=["Datum"])
    df["Arbetade timmar"] = df["Arbetade timmar"].astype(str).str.replace(",", ".")
    df["Arbetade timmar"] = pd.to_numeric(df["Arbetade timmar"], errors="coerce").fillna(0.0)

    daily_totals = (
        df.groupby(["Medarbetare", "Namn", "Datum"])["Arbetade timmar"]
        .sum()
        .reset_index()
    )
    return daily_totals


def sum_daily_flex(file_path: Path) -> pd.DataFrame | None:
    pf = skip_to_header(file_path)
    if pf is None or pf.empty:
        return None
    
    pf.columns = pf.columns.astype(str).str.strip().str.replace("\ufeff", "")
    columns_to_sum = ["ARB TID", "Flex +"]
    missing_cols = [col for col in columns_to_sum if col not in pf.columns]
    if missing_cols:
        print(f" \n  [Skipped]: '{file_path.name}' Can't identify required columns: {missing_cols}")
        return None
    for col in columns_to_sum:
        pf[col] = pf[col].astype(str).str.replace(",", ".")
        pf[col] = pd.to_numeric(pf[col], errors='coerce').fillna(0.0)

    pf["Total_Flex"] = pf["ARB TID"] + pf["Flex +"]
    split_date = pf["Datum"].astype(str).str.strip().str[:10]
    pf["Datum_Clean"] = pd.to_datetime(split_date, format="%Y-%m-%d", errors="coerce").dt.strftime("%Y-%m-%d")
    pf = pf.dropna(subset=["Datum_Clean"])
    daily_times = pf.groupby("Datum_Clean")["Total_Flex"].sum().reset_index()
    daily_times.rename(columns={"Datum_Clean": "Datum"}, inplace=True)
    return daily_times


def compare_employee_times(dept_df: pd.DataFrame, flex_df: pd.DataFrame, file_path: Path = None):
    """Merges department and flex times for a specific employee."""
    if dept_df.empty or flex_df is None or flex_df.empty:
        return None

    dept_df["Medarbetare"] = dept_df["Medarbetare"].astype(str)
    target = None

    if file_path:
        filename = file_path.name       
        num_match = re.search(r"(?<!\d)(\d{4})(?!\d)", filename) 
        if num_match and num_match.group(1) in dept_df["Medarbetare"].str.slice(0, 4).values:
            target = num_match.group(1)
        else:
            normalized_filename = re.sub(r"[_\-.]+", " ", filename).lower()
            for name in dept_df["Namn"].unique():
                if str(name).lower() in normalized_filename:
                    target = name
                    break

    while True:
        if not target:
            print(f"{'-'*60}\n Förhandsvisning av: {file_path.name}")
            pf = skip_to_header(file_path)
            if pf is not None:
                print(pf.head(5))
            
            filename_str = file_path.name if file_path else 'file'
            prompt_val = input(
                f"{'-' * 60}\n"
                f" Kunde ej matcha '{filename_str}'.\n"
                f" - Enter   : Hoppa över fil\n"
                f" - Ctrl+C : Avbryt\n"
                f"{'-' * 60}\n"
                f" Ange medarbetares namn eller ID: "
            ).strip()
            
            if not prompt_val:
                return None  
            target = prompt_val
            
        target_str = str(target).lower()
        person_dept_df = dept_df[
            (dept_df["Medarbetare"].str.slice(0, 4) == target_str) | (dept_df["Namn"].str.lower() == target_str)
        ].copy()

        if not person_dept_df.empty:
            break  

        print(f"  Hittar ej medarbetare '{target}' i databasen.")
        target = None

    emp_id = person_dept_df["Medarbetare"].dropna().iloc[0]
    emp_name = person_dept_df["Namn"].dropna().iloc[0]

    # Outer merge to keep all recorded days
#    comparison_df = pd.merge(person_dept_df[["Datum", "Arbetade timmar"]], flex_df, on="Datum", how="outer")
    comparison_df = pd.merge(person_dept_df[["Datum", "Arbetade timmar"]], flex_df, on="Datum", how="inner")
    
    comparison_df["Medarbetare"] = emp_id
    comparison_df["Namn"] = emp_name
    comparison_df["Arbetade timmar"] = comparison_df["Arbetade timmar"].fillna(0)
    comparison_df["Total_Flex"] = comparison_df["Total_Flex"].fillna(0)
    comparison_df["Diff"] = comparison_df["Arbetade timmar"] - comparison_df["Total_Flex"]

    comparison_df.sort_values("Datum", inplace=True)
    comparison_df.rename(columns={"Arbetade timmar": "EAM", "Total_Flex": "Stämplad Tid"}, inplace=True)

    cols = ["Medarbetare", "Namn", "Datum", "EAM", "Stämplad Tid", "Diff"]
    comparison_df = comparison_df[cols]

    return comparison_df, emp_name


def parse_temp_files(temp_dir: Path = Path("./temp_data"), output_dir: Path = Path("./output_data")):
    output_dir.mkdir(parents=True, exist_ok=True)
    
    files = [f for f in temp_dir.iterdir() if f.is_file() and f.suffix.lower() == ".csv"]
    eam_files = [f for f in files if "eam" in f.name.lower()]
    flex_files = [f for f in files if "eam" not in f.name.lower()]

    if not eam_files:
        print(" Ingen EAM-fil hittades.")
        return

    dept_df = parse_job_times(eam_files[0])
    for file_path in flex_files:
        flex_df = sum_daily_flex(file_path)
        
        result = compare_employee_times(dept_df, flex_df, file_path=file_path)
        if result is not None:
            result_df, emp_name = result
            new_emp_name = str(emp_name).replace(" ", "_").replace("/", "_")
            
            numeric_cols = ["EAM", "Stämplad Tid", "Diff"]
            for col in numeric_cols:
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce").round(2)
                    
            out_path_csv = output_dir / f"{new_emp_name}_jämförelse.csv"
            out_path_xlsx = output_dir / f"{new_emp_name}_jämförelse.xlsx"
            result_df.to_csv(out_path_csv, index=False, float_format="%.2f", encoding="utf-8-sig")
            
            excel_df = result_df.copy()
            excel_df["Datum"] = pd.to_datetime(excel_df["Datum"]).dt.date
            with pd.ExcelWriter(out_path_xlsx, engine="openpyxl") as writer:
                excel_df.to_excel(writer, index=False, sheet_name="Jämförelse")
                worksheet = writer.sheets["Jämförelse"]
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            print(f" Sparat jämförelse till: {out_path_xlsx}")
        else:
            print(f" Hoppade över fil: {file_path.name}")


def main():
    if len(sys.argv) > 1:
        files = [Path(p) for p in sys.argv[1:] if Path(p).exists()]
    else:
        files = select_files_via_gui()

    clear_temp_data()
    process_files(files)
    parse_temp_files()
    
    print("\n" + "=" * 60)
    print(" All operations failed successfully!")
    print("=" * 60)

if __name__ == "__main__":
    main()
