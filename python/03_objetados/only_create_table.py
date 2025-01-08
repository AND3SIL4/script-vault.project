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
        col_idx: int = int(params.get("col_idx"))
        cut_off_date: str = params.get("cut_off_date")

        ##Validate if all the required inputs are present
        if not all([file_path, sheet_name]):
            return "ERROR: a required input is missing"

        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        # Load data frame
        current_df = load_excel(file_path, sheet_name)

        # Set the months by name
        column_name = "MES_MOVIMIENTO"
        current_df = set_month_names(current_df, col_idx, column_name)

        # Fix white spaces
        current_df[column_name] = (
            current_df[column_name].astype(str).apply(clean_white_spaces)
        )

        # Generate the sum pivot table for the current month
        save_final_table(
            current_df, file_path, "TOTAL_REGISTROS", "count", column_name
        )
        return True, "Tabla generada correctamente en archivo de salida"
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

def save_final_table(
    df: pd.DataFrame, file_path: str, sheet_name: str, aggfunc: str, column_name: str
) -> None:
    """Save the final pivot table to the Excel file."""
    df[column_name] = df[column_name].astype(str).apply(clean_white_spaces)
    final_table = create_pivot_table(df, "VALOR RESERVA", aggfunc, column_name)
    add_total(final_table)
    sorted_table = sort_month_columns(final_table)
    save_to_file(sorted_table, file_path, sheet_name)

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
    """Add a total row to the pivot table."""
    current_sum = df.sum().astype(int)
    df.loc["TOTAL_ACTUAL"] = current_sum

def sort_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Sort the pivot table columns by month order."""
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
    """Remove white spaces from a string."""
    value = re.sub(r"[\s]", "", string)
    return value

def save_to_file(data_frame: pd.DataFrame, file_path: str, sheet_name: str) -> None:
    """Function to save the DataFrame to an Excel file."""
    with pd.ExcelWriter(
        file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        data_frame.to_excel(writer, sheet_name=sheet_name, index=True)

if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Temp\Objetados.xlsx",
        "sheet_name": "Objeciones 2022 - 2023 -2024",
        "col_idx": "44",
        "cut_off_date": "30/01/2024",
    }

    print(main(params))
