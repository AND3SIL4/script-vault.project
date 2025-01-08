import pandas as pd  # type: ignore
import os


def main(params: dict):
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        col_idx: int = int(params.get("col_idx"))
        inconsistencias_file: str = params.get("inconsistencias_file")
        sheet_name: str = "CASOS NUEVOS"
        list_file = params.get("list_file")

        ##Set initial variables
        if not all([file_path, col_idx, inconsistencias_file, sheet_name, list_file]):
            return "ERROR: An input required param is missing"

        ##Read book using pandas
        df: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )

        list_df: pd.DataFrame = pd.read_excel(
            list_file, sheet_name="COLUMNA CREDITO", engine="openpyxl"
        )

        ##Filter information and validate if the current data if number type
        df["is_valid"] = df.apply(
            lambda row: is_valid(
                str(row.iloc[col_idx]),
                list_df.iloc[:, 1].dropna().astype(str).values.tolist(),
                list_df.iloc[:, 2].dropna().astype(str).values.tolist(),
                str(row.iloc[15]),
                str(row.iloc[6]),
                list_df.iloc[:, 0].dropna().values.tolist(),
            ),
            axis=1,
        )
        ##Add inconsistencies to a filtered data frame
        filtered_file = df[~df["is_valid"]].copy()

        print(filtered_file.iloc[:, col_idx])

        ##Return and store the result
        if filtered_file.empty:
            return "Validación correcta, no se encontraron inconsistencias"
        else:
            filtered_file["COORDENADAS"] = filtered_file.apply(
                lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}",
                axis=1,
            )

            new_sheet_name = "ValidacionCredito"
            if os.path.exists(inconsistencias_file):
                with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
                    if new_sheet_name in xls.sheet_names:
                        exiting_df = pd.read_excel(
                            xls, sheet_name=new_sheet_name, engine="openpyxl"
                        )
                        filtered_file = pd.concat(
                            [exiting_df, filtered_file], ignore_index=True
                        )

            with pd.ExcelWriter(
                inconsistencias_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
                return "Inconsistencias registradas correctamente"

    except Exception as e:
        return f"Error: {e}"


def is_valid(
    value: str,
    mandatory: list[str],
    alpha_numeric: list[str],
    tomador: str,
    poliza: str,
    list_exception: list[str],
) -> bool:
    """Validate the value based on the given rules"""
    ##Delete white spaces at the end and start of the string
    value = value.strip()
    ##Validate if the value is empty string
    if value == "":
        return False

    ##Validate if the value is number
    if value.isdigit():
        return True

    ##Validate NaN (Not a Number)
    if value.lower() == "nan":
        if (tomador not in mandatory) or (int(poliza) in list_exception):
            return True
        else:
            return False

    ##Validate if the value is alphanumeric
    if value.isalnum():
        if tomador not in alpha_numeric:
            return False
        elif tomador not in mandatory:
            return True
        else:
            return True

    return False


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "col_idx": "98",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE REPARTO.xlsx",
    }

    print(main(params))
