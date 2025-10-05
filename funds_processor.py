import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence, Set, Tuple

import pandas as pd

try:
    import xlsxwriter  # type: ignore
except ImportError as exc:
    raise SystemExit(
        "The 'xlsxwriter' package is required. Install it with 'pip install xlsxwriter' and rerun this script."
    ) from exc

ALLOWED_BUSINESS_UNITS: Dict[str, Tuple[str, str]] = {
    "FI-US": ("FI-US", "AMRS"),
    "FI-EMEA": ("Fi-EMEA", "EMEA"),
    "FI-GMC-ASIA": ("FI-GMC-Asia", "APAC"),
}
ALLOWED_REVIEW_STATUS = {"APPROVED", "SUBMITTED"}
REGION_ORDER = ["AMRS", "EMEA", "APAC"]
REVIEW_STATUS_ORDER = ["Approved", "Submitted"]
BUSINESS_UNIT_LOOKUP = {key.upper(): value for key, value in ALLOWED_BUSINESS_UNITS.items()}


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def sanitize_sheet_name(name: str) -> str:
    invalid = set('[]:*?/\\')
    cleaned = ''.join('_' if c in invalid else c for c in name).strip()
    if not cleaned:
        cleaned = "Sheet"
    return cleaned[:31]


def make_unique_sheet_name(name: str, used: Set[str]) -> str:
    base = sanitize_sheet_name(name)
    candidate = base
    counter = 2
    while candidate in used:
        suffix = f"_{counter}"
        candidate = (
            base[: 31 - len(suffix)] + suffix
            if len(base) + len(suffix) > 31
            else base + suffix
        )
        counter += 1
    used.add(candidate)
    return candidate


def auto_fit_columns(worksheet, df: pd.DataFrame) -> None:
    for idx, column in enumerate(df.columns):
        series = df[column].astype(str)
        lengths = [len(value) for value in series]
        max_len = max(lengths, default=0)
        max_len = max(max_len, len(str(column)))
        worksheet.set_column(idx, idx, min(max_len + 2, 60))


def join_ids_for_output(values: List[str]) -> str:
    seen: Set[str] = set()
    ordered: List[str] = []
    for value in values:
        value = (value or "").strip()
        if not value or value in seen:
            continue
        seen.add(value)
        ordered.append(value)
    return ",".join(ordered)


def save_excel_with_tables(
    path: Path, sheets: Sequence[Tuple[str, pd.DataFrame]]
) -> None:
    ensure_directory(path.parent)
    used_sheet_names: Set[str] = set()
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for idx, (sheet_name, df) in enumerate(sheets, start=1):
            safe_sheet = make_unique_sheet_name(sheet_name, used_sheet_names)
            df_to_write = df.copy().fillna("")
            df_to_write.to_excel(writer, sheet_name=safe_sheet, index=False)
            worksheet = writer.sheets[safe_sheet]
            if df_to_write.shape[1] == 0:
                continue
            last_row = len(df_to_write)
            last_col = df_to_write.shape[1] - 1
            table_name = f"tbl_{idx}"
            worksheet.add_table(
                0,
                0,
                last_row,
                last_col,
                {
                    "name": table_name,
                    "columns": [
                        {"header": column} for column in df_to_write.columns
                    ],
                },
            )
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, last_row, last_col)
            auto_fit_columns(worksheet, df_to_write)


class StatsTracker:
    def __init__(self) -> None:
        self.tables: Dict[str, pd.DataFrame] = {}

    def set_region_summary(
        self, totals: pd.Series, blanks: pd.Series
    ) -> None:
        total_map = {k: int(v) for k, v in totals.to_dict().items()}
        blank_map = {k: int(v) for k, v in blanks.to_dict().items()}
        rows = []
        for region in REGION_ORDER:
            rows.append(
                {
                    "Region": region,
                    "Total Funds": total_map.get(region, 0),
                    "Country of Risk Blank": blank_map.get(region, 0),
                }
            )
        self.tables["Region Summary"] = pd.DataFrame(rows)

    def set_credit_coverage(
        self,
        region_present: pd.Series,
        region_missing: pd.Series,
        review_present: pd.Series,
        review_missing: pd.Series,
    ) -> None:
        region_present_map = {k: int(v) for k, v in region_present.to_dict().items()}
        region_missing_map = {k: int(v) for k, v in region_missing.to_dict().items()}
        region_rows = []
        for region in REGION_ORDER:
            region_rows.append(
                {
                    "Region": region,
                    "In Credit Studio": region_present_map.get(region, 0),
                    "Missing Credit Studio": region_missing_map.get(region, 0),
                }
            )
        self.tables["Credit Studio Coverage (Region)"] = pd.DataFrame(region_rows)

        review_present_map = {k: int(v) for k, v in review_present.to_dict().items()}
        review_missing_map = {k: int(v) for k, v in review_missing.to_dict().items()}
        review_rows = []
        for status in REVIEW_STATUS_ORDER:
            review_rows.append(
                {
                    "Review Status": status,
                    "In Credit Studio": review_present_map.get(status, 0),
                    "Missing Credit Studio": review_missing_map.get(status, 0),
                }
            )
        self.tables["Credit Studio Coverage (Review)"] = pd.DataFrame(review_rows)

    def set_country_match(
        self,
        region_match: pd.Series,
        region_mismatch: pd.Series,
        review_match: pd.Series,
        review_mismatch: pd.Series,
    ) -> None:
        region_match_map = {k: int(v) for k, v in region_match.to_dict().items()}
        region_mismatch_map = {k: int(v) for k, v in region_mismatch.to_dict().items()}
        region_rows = []
        for region in REGION_ORDER:
            region_rows.append(
                {
                    "Region": region,
                    "Country Match": region_match_map.get(region, 0),
                    "Country Mismatch": region_mismatch_map.get(region, 0),
                }
            )
        self.tables["Country of Risk Match (Region)"] = pd.DataFrame(region_rows)

        review_match_map = {k: int(v) for k, v in review_match.to_dict().items()}
        review_mismatch_map = {
            k: int(v) for k, v in review_mismatch.to_dict().items()
        }
        review_rows = []
        for status in REVIEW_STATUS_ORDER:
            review_rows.append(
                {
                    "Review Status": status,
                    "Country Match": review_match_map.get(status, 0),
                    "Country Mismatch": review_mismatch_map.get(status, 0),
                }
            )
        self.tables["Country of Risk Match (Review)"] = pd.DataFrame(review_rows)

    def export(self, path: Path) -> None:
        save_excel_with_tables(
            path,
            [(name, df.copy()) for name, df in self.tables.items()],
        )


@dataclass
class Config:
    all_funds_csv: Path = Path(r"D:\path\to\AllFunds.csv")
    extracted_data_dir: Path = Path(r"D:\path\to\credit_studio_exports")
    keys_workbook: Path = Path(r"D:\path\to\keys.xlsx")
    batch_size: int = 600

    def __post_init__(self) -> None:
        self.all_funds_csv = Path(self.all_funds_csv)
        self.extracted_data_dir = Path(self.extracted_data_dir)
        self.keys_workbook = Path(self.keys_workbook)

    @property
    def cleaned_output(self) -> Path:
        return self.all_funds_csv.with_name(f"{self.all_funds_csv.stem}_cleaned.xlsx")

    @property
    def blank_country_output(self) -> Path:
        return self.all_funds_csv.with_name(
            f"{self.all_funds_csv.stem}_blank_country_by_region.xlsx"
        )

    @property
    def stats_output(self) -> Path:
        return self.all_funds_csv.with_name(
            f"{self.all_funds_csv.stem}_processing_stats.xlsx"
        )

    @property
    def combined_credit_output(self) -> Path:
        return self.all_funds_csv.with_name(
            f"{self.all_funds_csv.stem}_credit_studio_combined.xlsx"
        )

    @property
    def missing_copers_output(self) -> Path:
        return self.all_funds_csv.with_name(
            f"{self.all_funds_csv.stem}_missing_copers.xlsx"
        )

    @property
    def corrections_output(self) -> Path:
        return self.all_funds_csv.with_name(
            f"{self.all_funds_csv.stem}_country_of_risk_corrections.xlsx"
        )


def load_and_clean_all_funds(config: Config) -> pd.DataFrame:
    if not config.all_funds_csv.exists():
        raise FileNotFoundError(
            f"All Funds CSV not found at {config.all_funds_csv}. Update Config.all_funds_csv."
        )
    print(f"Loading All Funds data from {config.all_funds_csv}")
    df = pd.read_csv(
        config.all_funds_csv,
        skiprows=1,
        dtype=str,
        keep_default_na=False,
        low_memory=False,
    )
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.replace("", pd.NA)
    df = df.dropna(how="all")
    required_columns = [
        "Business Unit",
        "Review Status",
        "Fund CoPER",
        "Country of Risk",
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise KeyError(
            f"Missing expected columns in All Funds file: {', '.join(missing_columns)}"
        )
    initial_rows = len(df)
    df["Business Unit"] = df["Business Unit"].fillna("").astype(str).str.strip()
    df["Business Unit Key"] = df["Business Unit"].str.upper()
    df = df[df["Business Unit Key"].isin(BUSINESS_UNIT_LOOKUP)].copy()
    df["Business Unit"] = df["Business Unit Key"].map(
        lambda key: BUSINESS_UNIT_LOOKUP[key][0]
    )
    df["Region"] = df["Business Unit Key"].map(
        lambda key: BUSINESS_UNIT_LOOKUP[key][1]
    )
    df.drop(columns=["Business Unit Key"], inplace=True)
    df["Review Status"] = df["Review Status"].fillna("").astype(str).str.strip()
    df = df[df["Review Status"].str.upper().isin(ALLOWED_REVIEW_STATUS)].copy()
    df["Review Status"] = df["Review Status"].str.title()
    df["Fund CoPER"] = df["Fund CoPER"].fillna("").astype(str).str.strip()
    missing_fund_coper = df["Fund CoPER"] == ""
    if missing_fund_coper.any():
        missing_count = int(missing_fund_coper.sum())
        print(f"[Warning] {missing_count} rows removed because Fund CoPER was blank.")
        df = df[~missing_fund_coper].copy()
    df["Country of Risk"] = df["Country of Risk"].fillna("").astype(str).str.strip()
    if "Region" not in df.columns:
        raise KeyError("Failed to derive Region column.")
    bu_index = df.columns.get_loc("Business Unit") + 1
    region_series = df.pop("Region")
    df.insert(bu_index, "Region", region_series)
    df = df.reset_index(drop=True)
    cleaned_rows = len(df)
    print(
        f"All Funds cleaned: {initial_rows} rows -> {cleaned_rows} rows after filters."
    )
    return df


def create_cleaned_workbook(df: pd.DataFrame, config: Config) -> None:
    save_excel_with_tables(
        config.cleaned_output, [("All Funds (Cleaned)", df)]
    )
    print(f"Cleaned All Funds workbook saved to {config.cleaned_output}.")


def generate_blank_country_workbook(
    df: pd.DataFrame, config: Config
) -> pd.DataFrame:
    blank_mask = df["Country of Risk"].astype(str).str.strip() == ""
    blank_df = df[blank_mask].copy()
    sheets = []
    for region in REGION_ORDER:
        region_df = blank_df[blank_df["Region"] == region].copy()
        sheets.append((region, region_df))
    save_excel_with_tables(config.blank_country_output, sheets)
    if blank_df.empty:
        print(
            "No blank Country of Risk rows detected; workbook contains empty tables for traceability."
        )
    else:
        print(
            f"Blank Country of Risk workbook saved to {config.blank_country_output} "
            f"({len(blank_df)} rows)."
        )
    return blank_df


def deliver_coper_batches(df: pd.DataFrame, config: Config) -> None:
    coper_series = df["Fund CoPER"].astype(str).str.strip()
    unique_copers = coper_series.drop_duplicates().tolist()
    total_ids = len(unique_copers)
    if total_ids == 0:
        print("No Fund CoPER IDs available after cleaning.")
        return
    batch_size = max(1, config.batch_size)
    total_batches = math.ceil(total_ids / batch_size)
    print(
        f"Preparing {total_ids} Fund CoPER IDs in batches of up to {batch_size} "
        f"(total batches: {total_batches})."
    )
    for batch_index in range(total_batches):
        start = batch_index * batch_size
        end = min(start + batch_size, total_ids)
        batch = unique_copers[start:end]
        batch_text = ",".join(batch)
        print("\n" + "=" * 72)
        print(
            f"Batch {batch_index + 1} of {total_batches} "
            f"({end - start} IDs • {end}/{total_ids} processed)"
        )
        print("-" * 72)
        print(batch_text)
        print("-" * 72)
        input(
            "Copy the comma-delimited IDs above, trigger the exposure extract, "
            "then press Enter for the next batch..."
        )
    while True:
        confirmation = input(
            "Type 'done' once all batches have been submitted to Credit Studio: "
        ).strip().lower()
        if confirmation in {"done", "yes", "y"}:
            break
        print("Waiting for confirmation... type 'done' when ready.")


def read_credit_exports(extracted_dir: Path) -> Tuple[pd.DataFrame, List[Path]]:
    xlsx_files = sorted(extracted_dir.glob("*.xlsx"))
    frames: List[pd.DataFrame] = []
    all_columns: List[str] = []
    for file in xlsx_files:
        df = pd.read_excel(file, dtype=str)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        frames.append(df)
        for column in df.columns:
            if column not in all_columns:
                all_columns.append(column)
    if not frames:
        return pd.DataFrame(), xlsx_files
    aligned = [frame.reindex(columns=all_columns) for frame in frames]
    combined = pd.concat(aligned, ignore_index=True)
    combined = combined.fillna("")
    return combined, xlsx_files


def select_first_25_columns(df: pd.DataFrame) -> pd.DataFrame:
    columns = df.columns.tolist()
    if "PMA" in columns:
        idx = columns.index("PMA")
        return df.iloc[:, : idx + 1]
    return df.iloc[:, : min(25, len(columns))]


def combine_credit_exports(
    clean_df: pd.DataFrame, config: Config, stats: StatsTracker
) -> pd.DataFrame:
    config.extracted_data_dir.mkdir(parents=True, exist_ok=True)
    print(
        f"Looking for Credit Studio export files (*.xlsx) in {config.extracted_data_dir}"
    )
    while True:
        combined_df, files = read_credit_exports(config.extracted_data_dir)
        if not files:
            input(
                "No Credit Studio export files found. "
                "Add the exposure files to the directory and press Enter to retry..."
            )
            continue
        if "Coper ID" not in combined_df.columns:
            raise KeyError(
                "Credit Studio exports are missing the 'Coper ID' column."
            )
        print(f"Combining {len(files)} Credit Studio files...")
        combined_df["Coper ID"] = combined_df["Coper ID"].astype(str).str.strip()
        combined_df = combined_df[combined_df["Coper ID"] != ""].reset_index(drop=True)
        if combined_df.empty:
            input(
                "Combined Credit Studio dataset is empty after removing blank Coper IDs. "
                "Review the export files, then press Enter to retry..."
            )
            continue
        save_excel_with_tables(
            config.combined_credit_output,
            [("Credit Studio Combined", combined_df)],
        )
        print(
            f"Combined Credit Studio data saved to {config.combined_credit_output} "
            f"({len(combined_df)} rows)."
        )
        unique_credit = combined_df.drop_duplicates(subset=["Coper ID"]).copy()
        if "Country of Risk" not in unique_credit.columns:
            raise KeyError(
                "Credit Studio combined data is missing the 'Country of Risk' column."
            )
        unique_credit["Country of Risk"] = (
            unique_credit["Country of Risk"].fillna("").astype(str).str.strip()
        )
        coverage_df = clean_df[["Fund CoPER", "Region", "Review Status"]].copy()
        coverage_df["Fund CoPER"] = (
            coverage_df["Fund CoPER"].fillna("").astype(str).str.strip()
        )
        unique_credit_ids = set(unique_credit["Coper ID"])
        coverage_df["In Credit Studio"] = coverage_df["Fund CoPER"].isin(
            unique_credit_ids
        )
        region_present = coverage_df[coverage_df["In Credit Studio"]][
            "Region"
        ].value_counts()
        region_missing = coverage_df[~coverage_df["In Credit Studio"]][
            "Region"
        ].value_counts()
        review_present = coverage_df[coverage_df["In Credit Studio"]][
            "Review Status"
        ].value_counts()
        review_missing = coverage_df[~coverage_df["In Credit Studio"]][
            "Review Status"
        ].value_counts()
        stats.set_credit_coverage(
            region_present,
            region_missing,
            review_present,
            review_missing,
        )
        stats.export(config.stats_output)
        missing_mask = ~coverage_df["In Credit Studio"]
        missing_rows = clean_df.loc[missing_mask].copy()
        if missing_rows.empty:
            print("All Fund CoPER IDs are present in the Credit Studio combined data.")
            return unique_credit
        print(
            f"Credit Studio data is missing {len(missing_rows)} Fund CoPER IDs."
        )
        missing_ids = missing_rows["Fund CoPER"].astype(str).str.strip()
        missing_ids_text = join_ids_for_output(missing_ids.tolist())
        print("Missing Fund CoPER IDs (comma separated):")
        print(missing_ids_text)
        first_25 = select_first_25_columns(missing_rows)
        save_excel_with_tables(
            config.missing_copers_output,
            [("Missing Fund CoPER IDs", first_25)],
        )
        print(
            f"Missing Fund CoPER details saved to {config.missing_copers_output}."
        )
        response = input(
            "Add the missing exposures to the Credit Studio export directory and "
            "press Enter to recombine, or type 'skip' to continue anyway: "
        ).strip().lower()
        if response == "skip":
            print("Continuing with incomplete Credit Studio coverage.")
            return unique_credit


def load_keys_table(keys_path: Path) -> pd.DataFrame:
    if not keys_path.exists():
        raise FileNotFoundError(
            f"Keys workbook not found at {keys_path}. Update Config.keys_workbook."
        )
    excel = pd.ExcelFile(keys_path)
    for sheet_name in excel.sheet_names:
        df = excel.parse(sheet_name, dtype=str)
        if df.empty:
            continue
        normalized_cols = {
            col.strip().lower(): col for col in df.columns if isinstance(col, str)
        }
        if "all funds" in normalized_cols and "credit studio" in normalized_cols:
            subset = df[
                [normalized_cols["all funds"], normalized_cols["credit studio"]]
            ].copy()
            subset = subset.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            subset = subset.dropna(how="any")
            subset = subset[(subset.iloc[:, 0] != "") & (subset.iloc[:, 1] != "")]
            subset.columns = ["All Funds", "Credit Studio"]
            print(
                f"Loaded normalization keys from sheet '{sheet_name}' "
                f"({len(subset)} rows)."
            )
            return subset
    raise ValueError(
        "Could not find 'All Funds' and 'Credit Studio' columns in the keys workbook."
    )


def build_normalization_map(keys_df: pd.DataFrame) -> Dict[str, Set[str]]:
    mapping: Dict[str, Set[str]] = {}
    for all_value, credit_value in keys_df[["All Funds", "Credit Studio"]].itertuples(
        index=False
    ):
        all_norm = (all_value or "").strip().lower()
        credit_norm = (credit_value or "").strip().lower()
        if not all_norm or not credit_norm:
            continue
        mapping.setdefault(all_norm, set()).add(credit_norm)
    return mapping


def countries_equal(
    all_value: str, credit_value: str, mapping: Dict[str, Set[str]]
) -> bool:
    all_value = (all_value or "").strip()
    credit_value = (credit_value or "").strip()
    if not all_value or not credit_value:
        return False
    all_key = all_value.lower()
    credit_key = credit_value.lower()
    if all_key in mapping:
        return credit_key in mapping[all_key]
    return all_key == credit_key


def match_country_of_risk(
    clean_df: pd.DataFrame,
    credit_df: pd.DataFrame,
    config: Config,
    stats: StatsTracker,
) -> pd.DataFrame:
    keys_df = load_keys_table(config.keys_workbook)
    normalization_map = build_normalization_map(keys_df)
    non_blank_df = clean_df[clean_df["Country of Risk"].astype(str).str.strip() != ""].copy()
    if non_blank_df.empty:
        print("No rows with Country of Risk available for matching.")
        stats.set_country_match(
            pd.Series(dtype=int),
            pd.Series(dtype=int),
            pd.Series(dtype=int),
            pd.Series(dtype=int),
        )
        stats.export(config.stats_output)
        return pd.DataFrame()
    credit_lookup = credit_df[["Coper ID", "Country of Risk"]].copy()
    credit_lookup["Country of Risk"] = (
        credit_lookup["Country of Risk"].fillna("").astype(str).str.strip()
    )
    merged = non_blank_df.merge(
        credit_lookup,
        left_on="Fund CoPER",
        right_on="Coper ID",
        how="left",
        suffixes=("", " (Credit Studio)"),
    )
    merged.rename(
        columns={
            "Country of Risk": "Country of Risk (All Funds)",
            "Country of Risk (Credit Studio)": "Country of Risk (Credit Studio)",
        },
        inplace=True,
    )
    merged["Country of Risk (All Funds)"] = (
        merged["Country of Risk (All Funds)"].fillna("").astype(str).str.strip()
    )
    merged["Country of Risk (Credit Studio)"] = (
        merged["Country of Risk (Credit Studio)"].fillna("").astype(str).str.strip()
    )
    merged["Country Match"] = merged.apply(
        lambda row: countries_equal(
            row["Country of Risk (All Funds)"],
            row["Country of Risk (Credit Studio)"],
            normalization_map,
        ),
        axis=1,
    )
    region_match = merged[merged["Country Match"]]["Region"].value_counts()
    region_mismatch = merged[~merged["Country Match"]]["Region"].value_counts()
    review_match = merged[merged["Country Match"]]["Review Status"].value_counts()
    review_mismatch = merged[~merged["Country Match"]]["Review Status"].value_counts()
    stats.set_country_match(region_match, region_mismatch, review_match, review_mismatch)
    stats.export(config.stats_output)

    correction_rows = merged[
        (~merged["Country Match"]) & (merged["Country of Risk (Credit Studio)"] != "")
    ].copy()
    correction_groups = []
    if not correction_rows.empty:
        grouped = correction_rows.groupby("Region")["Fund CoPER"]
        for region, series in grouped:
            correction_groups.append(
                {
                    "Region": region,
                    "Fund CoPER IDs": join_ids_for_output(series.tolist()),
                }
            )
    corrections_df = pd.DataFrame(correction_groups)
    if corrections_df.empty:
        corrections_df = pd.DataFrame(columns=["Region", "Fund CoPER IDs"])
    save_excel_with_tables(
        config.corrections_output,
        [("Country of Risk Corrections", corrections_df)],
    )
    if corrections_df.empty:
        print("No Country of Risk mismatches detected that require corrections.")
    else:
        print(
            f"Country of Risk corrections saved to {config.corrections_output} "
            f"({len(corrections_df)} region rows)."
        )
    return merged


def main() -> None:
    config = Config()
    stats = StatsTracker()
    clean_df = load_and_clean_all_funds(config)
    create_cleaned_workbook(clean_df, config)
    blank_df = generate_blank_country_workbook(clean_df, config)
    total_counts = clean_df["Region"].value_counts()
    blank_counts = blank_df["Region"].value_counts()
    stats.set_region_summary(total_counts, blank_counts)
    stats.export(config.stats_output)
    print(f"Initial stats saved to {config.stats_output}.")
    deliver_coper_batches(clean_df, config)
    credit_unique_df = combine_credit_exports(clean_df, config, stats)
    match_country_of_risk(clean_df, credit_unique_df, config, stats)
    print("\nProcessing complete. Review the generated outputs:")
    outputs = [
        config.cleaned_output,
        config.blank_country_output,
        config.stats_output,
        config.combined_credit_output,
        config.missing_copers_output,
        config.corrections_output,
    ]
    for path in outputs:
        status = "available" if path.exists() else "not generated in this run"
        print(f" - {path} ({status})")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nProcess interrupted by user.", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"\n[Error] {exc}", file=sys.stderr)
        sys.exit(1)
