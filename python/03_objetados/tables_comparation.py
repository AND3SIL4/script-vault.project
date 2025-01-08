import pandas as pd  # type: ignore
from datetime import datetime
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
        initial_date = params.get("initial_date")

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

        # Set the months by name
        column_name = "MES_MOVIMIENTO"
        current_filtered = set_month_names(current_filtered, col_idx, column_name)
        latest_filtered = set_month_names(latest_filtered, col_idx, column_name)
        current_df = set_month_names(current_df, col_idx, column_name)

        # Fix white spaces
        current_filtered[column_name] = (
            current_filtered[column_name].astype(str).apply(clean_white_spaces)
        )
        latest_filtered[column_name] = (
            latest_filtered[column_name].astype(str).apply(clean_white_spaces)
        )

        # Generate both sum and count pivot tables
        current_count_table = create_pivot_table(
            current_filtered, "VALOR RESERVA", "count", column_name
        )
        latest_count_table = create_pivot_table(
            latest_filtered, "VALOR RESERVA", "count", column_name
        )

        # Validate count tables
        is_count_valid = validate_tables(current_count_table, latest_count_table)

        # Save or report inconsistencies based on validation results
        if is_count_valid:
            save_final_table(
                current_df, file_path, "TOTAL_REGISTROS", "count", column_name
            )
            return True, "Tabla generada correctamente en archivo de salida"
        else:
            if not is_count_valid:
                report_inconsistencies(
                    current_count_table,
                    inconsistencias_file,
                    "ConteoRegistrosVsMesAnterior",
                )
            return False, "Inconsistencias reportadas en archivo"
    except Exception as e:
        return f"ERROR: {e} {traceback.format_exc()}"


def set_month_names(
    data_frame: pd.DataFrame, col_idx: int, col_name: str
) -> pd.DataFrame:
    """Set the month names in Spanish."""
    months_names: dict[int, str] = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }

    data_frame.iloc[:, col_idx] = pd.to_datetime(
        data_frame.iloc[:, col_idx], format="%d/%m/%Y", errors="coerce"
    )

    # Convertir valores de la columna a nombres de meses
    data_frame[col_name] = data_frame.apply(
        lambda row: months_names.get(int(row.iloc[col_idx].month)), axis=1
    )
    return data_frame


def load_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    """Load an Excel file into a DataFrame."""
    return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")


def filter_data(df: pd.DataFrame, col_idx: int, start_date, end_date) -> pd.DataFrame:
    """Filter the data frame based on the date range."""
    return df[
        (df.iloc[:, col_idx] > start_date) & (df.iloc[:, col_idx] < end_date)
    ].copy()


def create_pivot_table(
    df: pd.DataFrame, value_column: str, aggfunc: str, column_name: str
) -> pd.DataFrame:
    """Create a pivot table for the given aggregation function (sum or count)."""
    return pd.pivot_table(
        df,
        values=value_column,
        index="RAMO",
        columns=column_name,
        aggfunc=aggfunc,
        fill_value=0,
    ).astype(int)


def add_total(df: pd.DataFrame) -> None:
    """Validate if the current table matches the latest table."""
    current_sum = df.sum().astype(int)
    ##Add validation to show into the pivot table
    df.loc["TOTAL_ACTUAL"] = current_sum


def validate_tables(current_table: pd.DataFrame, latest_table: pd.DataFrame) -> bool:
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
    df: pd.DataFrame, file_path: str, sheet_name: str, aggfunc: str, column_name: str
) -> None:
    """Save the final pivot table to the Excel file."""
    df[column_name] = df[column_name].astype(str).apply(clean_white_spaces)
    final_table = create_pivot_table(df, "VALOR RESERVA", aggfunc, column_name)
    add_total(final_table)
    sorted_table = sort_month_columns(final_table)
    save_to_file(sorted_table, file_path, sheet_name)


def report_inconsistencies(df: pd.DataFrame, file_path: str, sheet_name: str) -> None:
    """Report inconsistencies in the validation."""
    sorted_df = sort_month_columns(df)
    message = append_inconsistencias(file_path, sheet_name, sorted_df)
    print(message)


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
    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=False)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=True, sheet_name=new_sheet)
            return "Inconsistencias registradas correctamente"
        
    except Exception as e:
        return f"Error al registrar inconsistencias: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Temp\Objetados.xlsx",
        "sheet_name": "Objeciones 2022 - 2023 -2024",
        "latest_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Output\Historico Objetados\Objetados 102024.xlsx",
        "col_idx": "44",
        "cut_off_date": "31/10/2024",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Temp\InconsistenciasBaseObjetados.xlsx",
        "initial_date": "01/01/2024"
    }

    print(main(params))
