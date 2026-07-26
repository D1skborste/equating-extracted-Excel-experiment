import sys
import subprocess

# Try to automatically install external dependencies
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
    print(" Denna databasen jämförs sedan med stämplade tider från Personec. \n För bästa upplevelse bör fil eller Excel-Sheet innehålla medarbetarens namn eller anställningsnummer")
    print(" Kan inte medarbetare identifieras kommer användaren uppmanas mata in uppgiften. \n")
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
    skip_amount = 0

    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            for i, line in enumerate(f):
                if target_key in line:
                    skip_amount = i
                    break
        
            df = pd.read_csv(file_path, skiprows=skip_amount, sep=None, engine="python")
        
        if df.columns[0] == "Medarbetare" and "Unnamed" in str(df.columns[1]):
            df.rename(columns={df.columns[1]: "Namn"}, inplace=True)
        df = df.loc[:, ~(df.columns.str.contains(r"^Unnamed:") & df.isna().all())]
        df = df.dropna(how="all").reset_index(drop=True)
        return df

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
    df.to_csv(output_dir / output_filename, index=False, encoding="utf-8")
    
    
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

        if len(df.columns) > 1 and df.columns[0] == "Medarbetare" and pd.isna(df.columns[1]):    
            df = df.rename(columns={df.columns[1]: "Namn"})
        
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


def parse_job_times(file_path):
    df = skip_to_header(file_path)
    df["Medarbetare"] = df["Medarbetare"].ffill()
    df["Namn"] = df["Namn"].ffill()
    df = df.dropna(subset=["Datum", "Arbetsorder"])
    df["Datum"] = pd.to_datetime(df["Datum"])
    df["Arbetade timmar"] = pd.to_numeric(df["Arbetade timmar"])
    daily_totals = (
        df.groupby(["Medarbetare", "Namn","Datum"])["Arbetade timmar"]
        .sum()
        .reset_index()
    )
    return daily_totals


def sum_daily_flex(file_path):
    pf = skip_to_header(file_path)
    columns_to_sum = ["ARB TID", "Flex +", "Flex -"]
        
    missing_cols = [col for col in columns_to_sum if col not in pf.columns]
    if missing_cols:
        print(f" \n  [Skipped]: '{file_path.name}' Can't identify required columns: {missing_cols}")
        return

    pf[columns_to_sum] = pf[columns_to_sum].apply(pd.to_numeric, errors='coerce')
    pf["Total_Flex"] = (
    pf["ARB TID"].fillna(0) 
    + pf["Flex +"].fillna(0) 
    - pf["Flex -"].fillna(0)
    )
    split_date = pf["Datum"].astype(str).str.strip().str[:10]
    pf["Datum"] = pd.to_datetime(split_date, format="%Y-%m-%d", errors="coerce")
    daily_times = pf.groupby("Datum")["Total_Flex"].sum().reset_index()
    return daily_times

# Parses all temp files, matches identified employees, and outputs final comparisons.
def parse_temp_files(temp_dir: Path = Path("./temp_data"), output_dir: Path = Path("./output_data")):
    output_dir.mkdir(parents=True, exist_ok=True)
    
    files = [f for f in temp_dir.iterdir() if f.is_file()]
    eam_files = [f for f in files if "eam" in f.name.lower()]
    flex_files = [f for f in files if "eam" not in f.name.lower()]

    if not eam_files:
        print(" Ingen EAM-fil hittades.")
        return

    dept_df = parse_job_times(eam_files[0])
    for file_path in flex_files:
        flex_df = sum_daily_flex(file_path)
        
        result_df = compare_employee_times(dept_df, flex_df, file_path=file_path)
        if result_df is not None:
            result_df, emp_name = result_df
            emp_name = str(emp_name).replace(" ", "_").replace("/", "_")
            out_path = output_dir / f"{emp_name}_jämförelse.csv"
            result_df.to_csv(out_path, index=False, float_format="%.2f", encoding="utf-8")
            print(f" Sparat jämförelse till: {out_path}")
        else:
            print(f" Hoppade över fil: {file_path.name}")
            
                
def compare_employee_times(dept_df, flex_df, file_path=None):
    """Merges department and flex times for a specific employee.    """

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
                f" - Enter  : Hoppa över fil\n"
                f" - Ctrl+C : Avbryt\n"
                f"{'-' * 60}\n"
                f" Ange medarbetares namn eller ID: "
            ).strip()
            
            if not prompt_val:
                return None  
            target = prompt_val
            
        person_dept_df = pd.DataFrame()
        if target:
            target = target.lower()
            person_dept_df = dept_df[
                (dept_df["Medarbetare"].str.slice(0, 4) == target) | (dept_df["Namn"].str.lower() == target)
            ].copy()

        if not person_dept_df.empty:
            break  

        print(f"  Hittar ej  medarbetare '{target}' i databasen.")
        target = None  # Reset target so input() triggers on the next loop iteration
        
    emp_name = person_dept_df["Namn"].dropna().iloc[0]

    # Find the overlapping date window where both datasets have logged dates
    min_date = max(person_dept_df["Datum"].min(), flex_df["Datum"].min())
    max_date = min(person_dept_df["Datum"].max(), flex_df["Datum"].max())
    person_dept_df = person_dept_df[(person_dept_df["Datum"] >= min_date) & (person_dept_df["Datum"] <= max_date)]
    flex_df = flex_df[(flex_df["Datum"] >= min_date) & (flex_df["Datum"] <= max_date)].copy()

    comparison_df = pd.merge(person_dept_df, flex_df, on="Datum", how="outer")
    comparison_df["Arbetade timmar"] = comparison_df["Arbetade timmar"].fillna(0)
    comparison_df["Total_Flex"] = comparison_df["Total_Flex"].fillna(0)

    comparison_df["Diff"] = (
        comparison_df["Arbetade timmar"] - comparison_df["Total_Flex"]
    )
    comparison_df.sort_values("Datum", inplace=True)
    comparison_df.rename(columns={"Arbetade timmar": "EAM", "Total_Flex": "Stämplad Tid"}, inplace=True)
    
#    print(comparison_df)
    return comparison_df, emp_name



def main():
    # 1. Grab dropped files or open GUI picker
    if len(sys.argv) > 1:
        files = [Path(p) for p in sys.argv[1:] if Path(p).exists()]
    else:
        files = select_files_via_gui()

    # 2. RUN PROCESSING LOGIC
    clear_temp_data()
    process_files(files)
    parse_temp_files()
    
    print("\n" + "=" * 60)
    print(" All operations failed successfully!")
    print("=" * 60)



if __name__ == "__main__":
    main()