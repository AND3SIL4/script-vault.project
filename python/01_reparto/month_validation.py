import pandas as pd  # type: ignore


def main(params) -> tuple:
    try:
        ## Set variables
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        inconsistencias_file = params.get("inconsistencias_file")

        ## Validate if all the required variables are present
        if not all([file_path, sheet_name, inconsistencias_file]):
            raise Exception("One or more input variables are not present")

        ## Read the file into a DataFrame
        reparto_data_frame = pd.read_excel(
            file_path, engine="openpyxl", sheet_name=sheet_name
        )

        ## Create a dictionary with the names of each month by number
        months_dict = {
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

        ## Create a function to validate both columns the date and the month
        def validate_month(date: str, month: str) -> bool:
            date_parse = pd.to_datetime(date, format="%Y-%m-%d", errors="coerce")
            get_month = date_parse.month
            standard_month = months_dict.get(get_month)
            return month == standard_month

        ## Pass data to the function for validating
        reparto_data_frame["is_valid"] = reparto_data_frame.apply(
            lambda row: validate_month(row.iloc[24], row.iloc[25]), axis=1
        )

        ## Create a inconsistencies data frame
        inconsistencias_data_frame = reparto_data_frame[~reparto_data_frame["is_valid"]]

        if not inconsistencias_data_frame.empty:
            ## Append the inconsistencies to the existing inconsistencias file
            return append_inconsistencies(
                inconsistencias_file, "MesAsignación", inconsistencias_data_frame
            )
        else:
            return (True, "INFO: no hay inconsistencias para registrar")

    except Exception as e:
        return (True, f"ERROR: {e}")


def append_inconsistencies(
    file_path: str, new_sheet: str, data_frame: pd.DataFrame
) -> str:
    with pd.ExcelFile(file_path, engine="openpyxl") as xls:
        if new_sheet in xls.sheet_names:
            existing = pd.read_excel(xls, engine="openpyxl", sheet_name=new_sheet)
            data_frame = pd.concat([existing, data_frame], ignore_index=True)

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        data_frame.to_excel(writer, sheet_name=new_sheet, index=False)
        return (True, "SUCCESS: inconsistencias registradas correctamente")


def get_excel_column_name(n: int) -> str:
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\INCONSISTENCIAS\InconBaseReparto.xlsx",
    }

    print(main(params))
