import pandas as pd  # type: ignore
import re
import traceback
import os


def main(params: dict):
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        latest_file: str = params.get("latest_file")
        col_idx: int = int(params.get("col_idx"))
        cut_off_date: str = params.get("cut_off_date")
        inconsistencias_file: str = params.get("inconsistencias_file")
        initial_date: str = params.get("initial_date")

        ##Validate if all the required inputs are present
        if not all([file_path, sheet_name, latest_file, inconsistencias_file]):
            return "ERROR: a required input is missing"

        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")
        cut_date = cut_off_date - pd.DateOffset(months=1)
        initial_date = pd.to_datetime(initial_date, format="%d/%m/%Y")

        # Load data frames
        current_df = load_excel(file_path, sheet_name)
        latest_df = load_excel(latest_file, sheet_name)
        # Filter data
        current_filtered = filter_data(current_df, col_idx, initial_date, cut_date)
        latest_filtered = filter_data(latest_df, col_idx, initial_date, cut_date)

        # Fix white spaces
        current_filtered["MES DE ASIGNACION"] = (
            current_filtered["MES DE ASIGNACION"].astype(str).apply(clean_white_spaces)
        )
        latest_filtered["MES DE ASIGNACION"] = (
            latest_filtered["MES DE ASIGNACION"].astype(str).apply(clean_white_spaces)
        )

        # Generate both sum and count pivot tables
        current_sum_table = create_pivot_table(current_filtered, "VALOR RESERVA", "sum")
        latest_sum_table = create_pivot_table(latest_filtered, "VALOR RESERVA", "sum")

        current_count_table = create_pivot_table(
            current_filtered, "VALOR RESERVA", "count"
        )
        latest_count_table = create_pivot_table(
            latest_filtered, "VALOR RESERVA", "count"
        )

        # Validate sum and count tables
        is_sum_valid = validate_tables(current_sum_table, latest_sum_table)
        is_count_valid = validate_tables(current_count_table, latest_count_table)

        # Save or report inconsistencies based on validation results
        if is_sum_valid:
            # Validate sum and count tables

            save_final_table(current_df, file_path, "VALOR_RESERVA", "sum")
            save_final_table(current_df, file_path, "TOTAL_REGISTROS", "count")
            return True
        else:
            if not is_sum_valid:
                report_inconsistencies(
                    current_sum_table,
                    inconsistencias_file,
                    "ValorReservaSumaInconsistencias",
                )
            if not is_count_valid:
                report_inconsistencies(
                    current_count_table,
                    inconsistencias_file,
                    "ValorReservaConteoInconsistencias",
                )
            return False
    except Exception as e:
        return f"ERROR: {e} {traceback.format_exc()}"


def load_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    """Load an Excel file into a DataFrame."""
    return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")


def filter_data(df: pd.DataFrame, col_idx: int, start_date, end_date) -> pd.DataFrame:
    """Filter the data frame based on the date range."""
    return df[
        (df.iloc[:, col_idx] > start_date) & (df.iloc[:, col_idx] < end_date)
    ].copy()


def create_pivot_table(
    df: pd.DataFrame, value_column: str, aggfunc: str
) -> pd.DataFrame:
    """Create a pivot table for the given aggregation function (sum or count)."""
    return pd.pivot_table(
        df,
        values=value_column,
        index="RAMO",
        columns="MES DE ASIGNACION",
        aggfunc=aggfunc,
        fill_value=0,
    ).astype(int)


def add_total(df: pd.DataFrame) -> None:
    """Validate if the current table matches the latest table."""
    current_sum = df.sum().astype(int)
    ##Add validation to show into the pivot table
    df.loc["TOTAL_ACTUAL"] = current_sum


def validate_tables(current_table: pd.DataFrame, latest_table: pd.DataFrame) -> bool:
    #!Debug
    # print(current_table)
    # print(latest_table)
    """Validate if the current table matches the latest table."""
    current_sum = current_table.sum().astype(int)
    latest_sum = latest_table.sum().astype(int)

    ##Add validation to show into the pivot table
    current_table.loc["TOTAL_ACTUAL"] = current_sum
    current_table.loc["TOTAL_ANTERIOR"] = latest_sum
    current_table.loc["VALIDACION"] = latest_sum == current_sum

    ##Return validation
    return current_sum.equals(latest_sum)


def save_final_table(
    df: pd.DataFrame, file_path: str, sheet_name: str, aggfunc: str
) -> None:
    """Save the final pivot table to the Excel file."""
    df["MES DE ASIGNACION"] = (
        df["MES DE ASIGNACION"].astype(str).apply(clean_white_spaces)
    )
    final_table = create_pivot_table(df, "VALOR RESERVA", aggfunc)
    add_total(final_table)
    sorted_table = sort_month_columns(final_table)
    save_to_file(sorted_table, file_path, sheet_name)


def report_inconsistencies(df: pd.DataFrame, file_path: str, sheet_name: str) -> None:
    """Report inconsistencies in the validation."""
    sorted_df = sort_month_columns(df)
    append_inconsistencias(file_path, sheet_name, sorted_df)


def sort_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    months = [
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE",
    ]
    columns_present = [mes for mes in months if mes in df.columns]
    sorted_df = df[columns_present]
    return sorted_df


def clean_white_spaces(string: str):
    value = re.sub(r"[\s]", "", string)
    return value


def save_to_file(data_frame: pd.DataFrame, file_path: str, sheet_name: str) -> None:
    """Function to save the DataFrame to an Excel file"""
    with pd.ExcelWriter(
        file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        data_frame.to_excel(writer, sheet_name=sheet_name, index=True)
        return "Tabla guardada correctamente"


def append_inconsistencias(file_path: str, new_sheet: str, data_frame) -> None:
    """This function get the inconsistencies data frame and append it into the inconsistencies file"""
    if os.path.exists(file_path):
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=False)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=True, sheet_name=new_sheet)
            return "Inconsistencias registradas correctamente"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "latest_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\BASE REPARTO\BASE DE REPARTO 112024.xlsx",
        "col_idx": "24",
        "cut_off_date": "30/12/2024",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "initial_date": "01/01/2024",
    }

    print(main(params))
