import pandas as pd
import os

def main(params: dict):
    try:
        # Set initial variables
        file_path = params.get("file_path")
        col_idx = int(params.get("col_idx"))
        inconsistencias_file = params.get("inconsistencias_file")
        sheet_name = params.get("sheet_name")

        ##Validate if all input params are present
        if not all([file_path, inconsistencias_file, sheet_name]):
           return "Error: an input params is missing"

        # Read book using pandas
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

        # Filter information and validate if the current data is number type
        df["is_number"] = df.iloc[:, col_idx].apply(lambda x: is_number(x))

        # Add inconsistencies to a filtered data frame
        filtered_file = df[~df["is_number"]].copy()

        # Return and store the result
        if filtered_file.empty:
            return "Validación correcta, no se encontraron inconsistencias"
        else:
            # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
            filtered_file['COORDENADAS'] = filtered_file.apply(
                lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}", axis=1
            )

            new_sheet_name = "ValidacionesTipoNumero"
            if os.path.exists(inconsistencias_file):
                with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
                    if new_sheet_name in xls.sheet_names:
                        existing_df = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
                        filtered_file = pd.concat([existing_df, filtered_file], ignore_index=True)

            with pd.ExcelWriter(inconsistencias_file, engine="openpyxl",mode="a",if_sheet_exists="replace",) as writer:
                filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
                return "Inconsistencias registradas correctamente"

    except Exception as e:
        return f"Error: {e}"


def is_number(value):
  try:
    ##Validate if the current value is null and return True
    if (str(value).lower() == "nan"):
      return True
    ##Try to convert the current value to know if is a number
    int(value)
    return True
  except:
    return False


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result