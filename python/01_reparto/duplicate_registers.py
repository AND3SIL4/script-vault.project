import pandas as pd  # type: ignore
import os


def main(params: dict) -> str:
    try:
        # Set the initial variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        exception_file: str = params.get("exception_file")
        sheet_exception: str = params.get("sheet_exception")

        # Validate if all the required inputs are present
        if not all(
            [
                file_path,
                sheet_name,
                inconsistencies_file,
                exception_file,
                sheet_exception,
            ]
        ):
            return "ERROR: an input required param is missing"

        # Read the work books
        base: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        exception_df = pd.read_excel(
            exception_file, sheet_name=sheet_exception, engine="openpyxl"
        )

        # Replace the NaN values with 0 in column "Credito"
        base.iloc[:, 98] = base.iloc[:, 98].fillna(0)

        # Create imaginary key to make validation
        base["KEY_1"] = base.iloc[:, 0].astype(str) + "-" + base.iloc[:, 2].astype(str)
        base["KEY_2"] = base["KEY_1"] + "-" + base.iloc[:, 32].astype(str)
        base["KEY_3"] = base["KEY_2"] + "-" + base.iloc[:, 34].astype(str)
        base["KEY_4"] = base["KEY_2"] + "-" + base.iloc[:, 27].astype(str)
        base["KEY_5"] = base["KEY_2"] + "-" + base.iloc[:, 98].astype(str)
        base["KEY_6"] = (
            base.iloc[:, 18].astype(str)
            + "-"
            + base.iloc[:, 32].astype(str)
            + "-"
            + base.iloc[:, 34].astype(str)
        )
        base["KEY_7"] = (
            base.iloc[:, 18].astype(str)
            + "-"
            + base.iloc[:, 32].astype(str)
            + "-"
            + base.iloc[:, 98].astype(str)
        )

        # Keys validation
        validation = base[
            (base["KEY_1"].duplicated(keep=False))
            & (base["KEY_2"].duplicated(keep=False))
            & (base["KEY_3"].duplicated(keep=False))
            & (base["KEY_4"].duplicated(keep=False))
            & (base["KEY_5"].duplicated(keep=False))
            & (base["KEY_6"].duplicated(keep=False))
            & (base["KEY_7"].duplicated(keep=False))
        ]

        inconsistencies = validation[
            (~validation["KEY_1"].isin(exception_df["KEY_1"]))
            & (~validation["KEY_2"].isin(exception_df["KEY_2"]))
            & (~validation["KEY_3"].isin(exception_df["KEY_3"]))
            & (~validation["KEY_4"].isin(exception_df["KEY_4"]))
            & (~validation["KEY_5"].isin(exception_df["KEY_5"]))
            & (~validation["KEY_6"].isin(exception_df["KEY_6"]))
            & (~validation["KEY_7"].isin(exception_df["KEY_7"]))
        ]
        print(inconsistencies)  #!Comment
        # Append inconsistencies into the file
        return validate_empty_df(
            inconsistencies_file, "ValidacionLlaves", inconsistencies
        )

    except Exception as e:
        return f"ERROR: {e}"


def append_inconsistencies_file(
    path_file: str, new_sheet: str, data_frame: pd.DataFrame
) -> str:
    """
    This function append a data frame into the inconsistencies file when is necessary

    Params:
    str: path_file (the path of the inconsistencies file)
    str: new_sheet (the name that you wanna put into the sheet in inconsistencies file)
    pd.DataFrame: data_frame (the data frame filtered previously)

    Returns:
    str: The confirm message
    """

    data_frame = data_frame.copy()
    if os.path.exists(path_file):
        data_frame["COORDENADAS_1"] = data_frame.apply(
            lambda row: f"{get_excel_column_name(0 + 1)}{row.name + 2}", axis=1
        )
        data_frame["COORDENADAS_2"] = data_frame.apply(
            lambda row: f"{get_excel_column_name(2 + 1)}{row.name + 2}", axis=1
        )

        with pd.ExcelFile(path_file, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=True)

        with pd.ExcelWriter(
            path_file, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, sheet_name=new_sheet, index=False)
            return "Inconsistencias registradas correctamente"


def validate_empty_df(path_file: str, new_sheet: str, data_frame: pd.DataFrame) -> str:
    """
    This function return a string message depends on the validation made

    Params:
    str: path_file (the path of the inconsistencies file)
    str: new_sheet (the name that you wanna put into the sheet in inconsistencies file)
    pd.DataFrame: data_frame (the data frame filtered previously)

    Return:
    str: confirmation message
    """
    if not data_frame.empty:
        return append_inconsistencies_file(path_file, new_sheet, data_frame)
    else:
        return "Validation realizada, no se encontraron inconsistencias"


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BASE DE REPARTO 2025.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE REPARTO.xlsx",
        "sheet_exception": "EXCEPCIONES VALIDACION LLAVES",
    }

    print(main(params))
