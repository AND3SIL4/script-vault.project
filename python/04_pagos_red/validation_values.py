import pandas as pd  # type: ignore
import numpy as np  # type: ignore
from typing import Optional
from datetime import datetime
from openpyxl import load_workbook  # type: ignore


class ValuesValidation:
    def __init__(
        self,
        file_path: str,
        inconsistencies_file: str,
        exception_file: str,
        sheet_name,
        file_name: str,
        previous_file: str,
        temp_file: str,
        historic_file: str,
    ):
        self.file_path = file_path
        self.inconsistencies_file = inconsistencies_file
        self.exception_file = exception_file
        self.sheet_name = sheet_name
        self.file_name = file_name
        self.previous_file = previous_file
        self.temp_file = temp_file
        self.historic_file = historic_file

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        return pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl", dtype=str
        )

    def save_inconsistencies_file(self, df: pd.DataFrame, new_sheet: str) -> bool:
        """Method to save the inconsistencies data frame into the inconsistencies file"""
        try:
            with pd.ExcelFile(self.inconsistencies_file, engine="openpyxl") as xls:
                if new_sheet in xls.sheet_names:
                    existing = pd.read_excel(
                        xls, engine="openpyxl", sheet_name=new_sheet
                    )
                    df = pd.concat([existing, df], ignore_index=True)

            with pd.ExcelWriter(
                self.inconsistencies_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                df.to_excel(writer, sheet_name=new_sheet, index=False)
                return True
        except Exception as e:
            print(f"Error: {e}")
            return False

    def excel_col_name(self, number) -> str:
        """Method to convert (1-based) to Excel column name"""
        result = ""
        while number > 0:
            number, reminder = divmod(number - 1, 26)
            result = chr(65 + reminder) + result
        return result

    def validate_inconsistencies(
        self, df: pd.DataFrame, col_idx, sheet_name: str
    ) -> str:
        """Method to validate the inconsistencies before append in a inconsistencies file"""
        if not df.empty:
            df = df.copy()
            if isinstance(col_idx, int):
                df[f"COORDENADAS"] = df.apply(
                    lambda row: f"{self.excel_col_name(col_idx + 1)}{row.name + 2}",
                    axis=1,
                )
            else:
                for i in col_idx:
                    df[f"COORDENADAS_{i + 2}"] = df.apply(
                        lambda row: f"{self.excel_col_name(i+1)}{row.name + 2}",
                        axis=1,
                    )
            self.save_inconsistencies_file(df, sheet_name)
            return "Success: Inconsistencies guardadas correctamente"
        else:
            return "Info: Validacion realizada, no se encontraron inconsistencias"

    def get_file_date(self) -> str:
        """Method to get the file date"""
        start: str = self.file_name.find("(") + 1
        end: str = self.file_name.find(")")

        # Extract the text
        if start > 0 and end > 0:
            file_date = self.file_name[start:end]
            file_date = file_date.replace("-", "/")
            return file_date
        else:
            return "Error obteniendo fecha"


# Instance the main class
values_validation: Optional[ValuesValidation] = None


def extract_data_from_propuesta(data_frame: pd.DataFrame) -> pd.DataFrame:
    """Method to get the importan data from propuesta de pagos file and the return it into a data frame"""
    # Filter data frame by number of index only with the important columns
    # N° radicado casa matriz = 2
    # Valor movimiento 100% = 45
    data_frame = data_frame.iloc[:, [2, 45]].copy()
    # Drop rows with NaN values in the first column (N° radicado casa matriz)
    data_frame.dropna(subset=[data_frame.columns[0]], inplace=True)
    # Add additional columns
    data_frame["LLAVE RADICADO + MOVIMIENTO"] = (
        data_frame[data_frame.columns[0]] + " " + data_frame[data_frame.columns[1]]
    )
    data_frame["AÑO"] = datetime.now().year
    data_frame.iloc[:, 0] = data_frame.iloc[:, 0]
    return data_frame


def cross_file(propuesta_df: pd.DataFrame, acm_df: pd.DataFrame) -> pd.DataFrame:
    """Method to cross the propuesta and acm report and return the merged data frame"""
    # Merge the two data frames based on the radicado number
    merged_df = pd.merge(
        propuesta_df,
        acm_df,
        left_on=propuesta_df.columns[0],
        right_on=acm_df.columns[0],
        how="left",
        suffixes=("_PROPUESTA", "_ACM"),
    )
    # Filter file before return it
    merged_df = merged_df[
        [
            "AÑO",
            "No DE RADICADO CASA MATRIZ",
            "VR. MOVIMIENTO 100%",
            "LLAVE RADICADO + MOVIMIENTO",
            "Valor Liquidado",
            "valor aprobado",
        ]
    ]
    return merged_df


def apply_formulas(data_frame: pd.DataFrame, historical_df: pd.DataFrame) -> None:
    """Method to apply formulas to the merged data frame"""
    # Get the list of the historical report
    historical_radicados: list[str] = (
        historical_df.iloc[:, 1].dropna().astype(str).to_list()
    )
    historical_key: list[str] = historical_df.iloc[:, 3].dropna().astype(str).to_list()

    # Convertir columnas a valores numéricos
    columns_to_convert = ["Valor Liquidado", "valor aprobado", "VR. MOVIMIENTO 100%"]
    for column in columns_to_convert:
        data_frame[column] = (
            data_frame[column]
            .str.replace(",", "")
            .str.replace(".", "")
            .str.replace(" ", "")
        )
        data_frame[column] = pd.to_numeric(data_frame[column], errors="coerce").fillna(
            0
        )

    # Aplicar fórmulas
    data_frame[historical_df.columns[6]] = (
        data_frame["valor aprobado"] - data_frame["Valor Liquidado"]
    )

    # Validate valores reported
    data_frame[historical_df.columns[7]] = data_frame.iloc[:, 2].astype(
        int
    ) == data_frame.iloc[:, 5].astype(int)

    # Validate records duplicates
    ## Radicado column
    add_number_of_duplicates(data_frame, historical_df, historical_radicados, 1, 8)

    ## Key column: (radicado number + concepto value)
    add_number_of_duplicates(data_frame, historical_df, historical_key, 3, 9)

    # Validate format values in columns
    data_frame[historical_df.columns[10]] = data_frame[data_frame.columns[1]].apply(
        lambda value: str(value).replace(",", "").replace(".", "").isdigit()
    )
    data_frame[historical_df.columns[11]] = data_frame[data_frame.columns[2]].apply(
        lambda value: str(value).replace(",", "").replace(".", "").isdigit()
    )

    data_frame[historical_df.columns[12]] = values_validation.get_file_date()

    # Validate if the "radicado" number starts with a different number of year
    # Get the current year
    year = datetime.now().year

    def ensure_sheet_exists(file_path: str, sheet_name: str) -> None:
        try:
            workbook = load_workbook(file_path)
            # Verifica si la hoja ya existe
            if sheet_name not in workbook.sheetnames:
                # Si la hoja no existe, crea una nueva
                new_sheet = workbook.create_sheet(sheet_name)
                # Verifica si la hoja anterior existe
                if (
                    sheet_name != "Sheet1"
                    and f"{int(sheet_name)-1}" in workbook.sheetnames
                ):
                    previous_sheet = workbook[f"{int(sheet_name)-1}"]
                    # Copiar los encabezados de la hoja anterior
                    for col in range(1, previous_sheet.max_column + 1):
                        new_sheet.cell(
                            row=1,
                            column=col,
                            value=previous_sheet.cell(row=1, column=col).value,
                        )

                workbook.save(file_path)
                print(f"La hoja '{sheet_name}' se creó correctamente.")
        except Exception as e:
            print(e)

    ensure_sheet_exists(values_validation.historic_file, str(year))

    # Sub function to validate if the "radicado" number starts with a different number aside current year
    def validate_previous_radicados(radicado: str) -> int:
        print("One")
        current_radicado_year = radicado[:4]
        if current_radicado_year != str(year):
            previous_dfs: list[pd.DataFrame] = []
            for year_to_read in range(int(current_radicado_year), year + 1):
                print("One one")
                df: pd.DataFrame = values_validation.read_excel(
                    values_validation.historic_file, sheet_name=str(year_to_read)
                )
                previous_dfs.append(df)
            entire_df: pd.DataFrame = pd.concat(previous_dfs, ignore_index=True)
            entire_list: list[str] = (
                entire_df[entire_df.columns[1]].dropna().astype(str).to_list()
            )
            return entire_list.count(radicado)

    def validate_previous_key(key: str) -> int:
        print("Two")
        current_key_year = key[:4]
        if current_key_year != str(year):
            previous_dfs: list[pd.DataFrame] = []
            for year_to_read in range(int(current_key_year), year + 1):
                print("Two two")
                df: pd.DataFrame = values_validation.read_excel(
                    values_validation.historic_file, sheet_name=str(year_to_read)
                )
                previous_dfs.append(df)
            entire_df: pd.DataFrame = pd.concat(previous_dfs, ignore_index=True)
            key_column_list: list[str] = (
                entire_df.iloc[:, 3].str.replace(" -", "").astype(str).to_list()
            )
            # Return the amount of values in the historical key list
            return key_column_list.count(key)

    data_frame["is_previous_key"] = data_frame.apply(
        lambda row: validate_previous_key(str(row.iloc[3])),
        axis=1,
    )
    data_frame["is_previous_radicado"] = data_frame.apply(
        lambda row: validate_previous_radicados(str(row.iloc[1])),
        axis=1,
    )

    # Sum the previous radicado in previous years
    data_frame[historical_df.columns[8]] = data_frame[historical_df.columns[8]].astype(
        int
    ) + data_frame["is_previous_radicado"].fillna(0).astype(int)

    data_frame[historical_df.columns[9]] = data_frame[historical_df.columns[9]].astype(
        int
    ) + data_frame["is_previous_key"].fillna(0).astype(int)

    # Ignore no needed temp columns
    data_frame = data_frame.iloc[:, :13]
    print(data_frame)  # <-- this is the final data frame
    # Set up unique columns
    data_frame.columns = historical_df.columns
    return data_frame


def add_number_of_duplicates(
    data_frame: pd.DataFrame,
    historical_df: pd.DataFrame,
    previous_list: list,
    index_col: int,
    historical_index_col: int,
) -> None:
    ## Key column: (radicado number + concepto value)
    key_col = data_frame.columns[index_col]
    key_count = data_frame[key_col].map(data_frame[key_col].value_counts())
    key_in_list = data_frame[key_col].map(lambda value: previous_list.count(value))
    data_frame[historical_df.columns[historical_index_col]] = key_count + key_in_list


def validate_values(acm_file: str) -> None:
    try:
        # Extract data from propuesta de pagos file
        propuesta_pago_df = values_validation.read_excel(
            values_validation.file_path, values_validation.sheet_name
        )
        # Get previous data frame
        historical_df: pd.DataFrame = values_validation.read_excel(
            values_validation.previous_file, values_validation.sheet_name
        )
        # Get the ACM report data fame
        acm_report: pd.DataFrame = values_validation.read_excel(
            acm_file, "FCT_RS_REPORTE_WS_AUDITORIA"
        )
        # Delete te no needed columns and rows
        acm_report = acm_report.iloc[3:, 1:]
        acm_report.columns = acm_report.iloc[0]
        acm_report = acm_report.iloc[1:].reset_index(drop=True)
        acm_report = acm_report[["id cuenta", "valor aprobado", "Valor Liquidado"]]

        # Validate and extract the important data from propuesta and acm report
        propuesta_df: pd.DataFrame = extract_data_from_propuesta(propuesta_pago_df)
        merged_df: pd.DataFrame = cross_file(propuesta_df, acm_report)
        # Apply formulas to validate inconsistencies
        filled_df: pd.DataFrame = apply_formulas(merged_df, historical_df)
        # Report inconsistencies
        report_inconsistencies(filled_df)
        # Save the final file into temp file folder
        filled_df.to_excel(
            values_validation.temp_file,
            sheet_name=values_validation.sheet_name,
            index=False,
        )
        return True, f"Function '{validate_values.__name__}' executed successfully"

    except Exception as e:
        return False, f"Error: {e}"


def report_inconsistencies(data_frame: pd.DataFrame) -> None:
    """Method to generate a report of inconsistencies and save it into tbe correct file"""
    try:
        # 1. Valores validation
        valores_exception_df: pd.DataFrame = values_validation.read_excel(
            values_validation.exception_file, "VALIDACION VALORES"
        )
        valores_exception_list: list[str] = (
            valores_exception_df.iloc[:, 0].dropna().astype(str).to_list()
        )
        valores_inconsistencies: pd.DataFrame = data_frame[~data_frame.iloc[:, 7]]
        valores_inconsistencies = valores_inconsistencies[
            ~valores_inconsistencies.iloc[:, 3].isin(valores_exception_list)
        ]
        # Save the inconsistencies
        values_validation.validate_inconsistencies(
            valores_inconsistencies, [3, 7], "ValidacionValores"
        )
        # 2. Radicados number duplicated
        duplicados_exception_list: pd.DataFrame = values_validation.read_excel(
            values_validation.exception_file, "VALIDACION DUPLICADOS"
        )
        radicados_exception_list: list[str] = (
            duplicados_exception_list.iloc[:, 0].dropna().astype(str).to_list()
        )
        radicados_inconsistencies: pd.DataFrame = data_frame[
            data_frame.iloc[:, 8].astype(int) > 1
        ]
        radicados_inconsistencies = radicados_inconsistencies[
            ~radicados_inconsistencies.iloc[:, 1].isin(radicados_exception_list)
        ]
        # Save the inconsistencies
        values_validation.validate_inconsistencies(
            radicados_inconsistencies, [1, 8], "ValidacionRadicadosDuplicados"
        )
        # 3. Key duplicated
        key_exception_list: list[str] = (
            duplicados_exception_list.iloc[:, 1].dropna().astype(str).to_list()
        )
        key_inconsistencies: pd.DataFrame = data_frame[
            data_frame.iloc[:, 9].astype(int) > 1
        ]
        key_inconsistencies = key_inconsistencies[
            ~key_inconsistencies.iloc[:, 3].isin(key_exception_list)
        ]
        # Save the inconsistencies
        values_validation.validate_inconsistencies(
            key_inconsistencies, [3, 9], "ValidacionKeyDuplicados"
        )
        # 4. Radicado format
        format_exception_df: pd.DataFrame = values_validation.read_excel(
            values_validation.exception_file, "VALIDACION FORMATOS"
        )
        radicados_format_exception_list: list[str] = (
            format_exception_df.iloc[:, 0].dropna().astype(str).to_list()
        )
        radicados_format_inconsistencies: pd.DataFrame = data_frame[
            ~data_frame.iloc[:, 10].astype(bool)
        ]
        radicados_format_inconsistencies = radicados_format_inconsistencies[
            ~radicados_format_inconsistencies.iloc[:, 3].isin(
                radicados_format_exception_list
            )
        ]
        # Save the inconsistencies
        values_validation.validate_inconsistencies(
            radicados_format_inconsistencies, [3, 10], "ValidacionRadicadoFormato"
        )
        # 5. Valor 100% format
        valor_100_format_exception_list: list[str] = (
            format_exception_df.iloc[:, 1].dropna().astype(str).to_list()
        )
        valor_100_format_inconsistencies: pd.DataFrame = data_frame[
            ~data_frame.iloc[:, 11].astype(bool)
        ]
        valor_100_format_inconsistencies = valor_100_format_inconsistencies[
            ~valor_100_format_inconsistencies.iloc[:, 2].isin(
                valor_100_format_exception_list
            )
        ]
        # Save the inconsistencies
        values_validation.validate_inconsistencies(
            valor_100_format_inconsistencies, [2, 11], "ValidacionValor100Formato"
        )
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False


# Call the main function
def main(params: dict) -> bool:
    try:
        global values_validation
        if values_validation is None:
            values_validation = ValuesValidation(
                file_path=params.get("file_path"),
                inconsistencies_file=params.get("inconsistencies_file"),
                exception_file=params.get("exception_file"),
                sheet_name=params.get("sheet_name"),
                file_name=params.get("file_name"),
                previous_file=params.get("previous_file"),
                temp_file=params.get("temp_file"),
                historic_file=params.get("historic_file"),
            )
            return True
    except Exception as e:
        print(f"Error: {e}")
        return False


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\PROPUESTA DE PAGO 1 Y 2  (02-01-2025).xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\InconsistenciasBasePagosRedAsistencial.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\EXCEPCIONES BASE PAGOS RED ASISTENCIAL.xlsx",
        "sheet_name": "Propuesta",
        "file_name": "PROPUESTA DE PAGO 1 Y 2  (02-01-2025).xlsx",
        "previous_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Output\Validacion valores\Validacion valores 04122024.xlsx",
        "temp_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\Pagos red asistencial 18122024.xlsx",
        "historic_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\VALIDADOR PAGOS.xlsx",
    }
    main(params)
    incomes = r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\FCT_RS_REPORTE_WS_AUDITORIA.xlsx"
    print(validate_values(incomes))
