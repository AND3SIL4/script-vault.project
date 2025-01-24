import pandas as pd  # type: ignore
from typing import Tuple
from datetime import datetime


def main(params: dict) -> Tuple[bool, str]:
    try:
        # Set initial local variables
        previous_year_file: str = params.get("previous_year_file")
        current_file: str = params.get("current_file")
        sheet_name: str = params.get("sheet_name")
        exception_file: str = params.get("exception_file")
        exception_sheet_name: str = params.get("exception_sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")

        # Load the previous year and current files
        previous_df = load_file(file_path=previous_year_file, sheet_name=sheet_name)
        current_df = load_file(file_path=current_file, sheet_name=sheet_name)
        except_df = load_file(file_path=exception_file, sheet_name=exception_sheet_name)

        # Create a new column with the year into the radicado number
        current_df["validate_radicado_year"] = current_df.iloc[:, 2].str[:4]

        current_year = datetime.now().year
        previous_year = current_year - 1  # Get the previous year subtracting 1

        # Filter current df by the previous year
        cur_filtered_df = current_df[
            current_df["validate_radicado_year"] == str(previous_year)
        ].copy()
        # Create keys with specific columns
        add_keys(cur_filtered_df)
        add_keys(previous_df)

        inconsistencies = cur_filtered_df[
            (cur_filtered_df["KEY_1"].isin(previous_df["KEY_1"]))
            & (cur_filtered_df["KEY_2"].isin(previous_df["KEY_2"]))
            & (cur_filtered_df["KEY_3"].isin(previous_df["KEY_3"]))
            & (cur_filtered_df["KEY_4"].isin(previous_df["KEY_4"]))
            & (cur_filtered_df["KEY_5"].isin(previous_df["KEY_5"]))
            & (cur_filtered_df["KEY_6"].isin(previous_df["KEY_6"]))
            & (cur_filtered_df["KEY_7"].isin(previous_df["KEY_7"]))
        ].copy()
        # Validate if there is any exception record
        # before report inconsistencies to the user
        # Filter out records that are in the exceptions DataFrame
        inconsistencies["is_exception"] = (
            inconsistencies["KEY_1"].isin(except_df["KEY_1"])
            & inconsistencies["KEY_2"].isin(except_df["KEY_2"])
            & inconsistencies["KEY_3"].isin(except_df["KEY_3"])
            & inconsistencies["KEY_4"].isin(except_df["KEY_4"])
            & inconsistencies["KEY_5"].isin(except_df["KEY_5"])
            & inconsistencies["KEY_6"].isin(except_df["KEY_6"])
            & inconsistencies["KEY_7"].isin(except_df["KEY_7"])
        )

        # Filter the inconsistencies by the exception records
        inconsistencies_validated: pd.DataFrame = inconsistencies[
            ~inconsistencies["is_exception"]
        ].copy()

        return save_inconsistencies(
            df=inconsistencies_validated,
            new_sheet="ResiduosDuplicadosAñoAnterior",
            inconsistencies_file=inconsistencies_file,
        )

    except Exception as e:
        return (False, str(e))


def add_keys(base: pd.DataFrame) -> None:
    base.iloc[:, 98] = base.iloc[:, 98].fillna(0)
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


def excel_col_name(number: int) -> str:
    result = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def save_inconsistencies(
    df: pd.DataFrame, new_sheet: str, inconsistencies_file: str
) -> str:
    """
    This function saves a DataFrame to an Excel file
    """
    if not df.empty:
        with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing_df = pd.read_excel(
                    xls, sheet_name=new_sheet, engine="openpyxl"
                )
                df = pd.concat([existing_df, df], ignore_index=True)

        with pd.ExcelWriter(
            inconsistencies_file, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=new_sheet, index=False)
            return "SUCCESS: inconsistencias registradas correctamente"
    else:
        return "INFO: No se encontraron inconsistencias para registrar"


def load_file(file_path: str, sheet_name) -> pd.DataFrame:
    """
    This function loads an Excel file into a DataFrame.
    """
    return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl", dtype=str)


if __name__ == "__main__":
    params = {
        "previous_year_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\BASE REPARTO\BASE DE REPARTO 122024.xlsx",
        "current_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BASE DE REPARTO 2025.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE REPARTO.xlsx",
        "exception_sheet_name": "EXCEPCIONES CAMBIO AÑO",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\INCONSISTENCIAS\InconBaseReparto.xlsx",
    }
    result = main(params)
    print(result)
