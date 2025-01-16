import pandas as pd  # type:ignore
from typing import Optional
import os
import re


class FirstValidationGroup:
    """Class to make the fist validation in the 'Base de Pagos' process"""

    def __init__(
        self,
        path_file: str,
        sheet_name: str,
        inconsistencies_file: str,
        exception_file: str,
    ):
        self.path_file = path_file
        self.sheet_name = sheet_name
        self.inconsistencies_file = inconsistencies_file
        self.exception_file = exception_file

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    def save_inconsistencies_file(self, df: pd.DataFrame, new_sheet: str) -> bool:
        if os.path.exists(self.inconsistencies_file):
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
        else:
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
            return "SUCCESS: Inconsistencies guardadas correctamente"
        else:
            return "INFO: Validacion realizada, no se encontraron inconsistencias"

    def validate_empty_col(self, col_idx: int, mandatory: bool) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        if mandatory:
            data_frame["is_valid"] = data_frame.iloc[:, col_idx].isna()
            inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
            return self.validate_inconsistencies(
                inconsistencies, col_idx, "ValidacionColumnasVacias"
            )
        else:
            data_frame["is_valid"] = data_frame.iloc[:, col_idx].isna()
            inconsistencies: pd.DataFrame = data_frame[data_frame["is_valid"]]
            return self.validate_inconsistencies(
                inconsistencies, col_idx, "ValidacionColumnasVacias"
            )

    def number_type(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(self.exception_file, "LISTAS")
        list_exception: list = exception_df["SAP"].dropna().astype(str).tolist()
        print(list_exception)

        def validate_with_exception_list(value: str) -> bool:
            value = value.replace(".", "")
            try:
                int(value)
                return True
            except ValueError:
                return value in list_exception

        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(lambda value: validate_with_exception_list(value))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, "DatoTipoNumero")

    def date_type(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = pd.to_datetime(
            data_frame.iloc[:, col_idx], errors="coerce"
        ).notna()
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, "DatosTipoFecha")

    def value_length(self, col_idx: int, length: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: len(str(value)) == length
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, "LongitudValor")

    def validate_exception_list(
        self,
        col_idx: int,
        exception_col_name: int,
        exception_sheet: str,
        new_sheet: str,
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_data_frame: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        col_exception: pd.Series = exception_data_frame[exception_col_name].dropna()
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].isin(col_exception)
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def no_special_characters(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(
                lambda value: not pd.isna(value)
                and bool(re.search(r"[^a-zA-Z0-9]", value))
            )
        )

        inconsistencies: pd.DataFrame = data_frame[data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionCaracteresEspaciales"
        )

    def month_depends_on_date(
        self, date_idx: int, month_idx: int, exception_sheet: str, exception_idx: int
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        exception_list: pd.Series = (
            exception_df.iloc[:, exception_idx].astype(str).dropna().to_list()
        )

        ## Create s sub function to know the correct month depends on the number
        months: dict = {
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

        ## Create a sub function to validate the consistency of the date
        def validate_consistency(date: str, month: str, radicado: str) -> bool:
            date_parse = pd.to_datetime(date, format="%Y-%m-%d", errors="coerce")
            get_month = date_parse.month
            ## Call the dictionary
            standard_month = months.get(get_month)
            return (month == standard_month) or (radicado in exception_list)

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_consistency(
                row.iloc[date_idx], str(row.iloc[month_idx]), str(row.iloc[2])
            ),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, month_idx, "ValidacionMesCorte"
        )

    def radicado_format(self, col_idx) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: bool(re.search(r"^\d{4}\s\d{2}\s\d{3}\s\d{6}$", str(value)))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "FormatoNumeroRadicado"
        )

    def acuerdo_range(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: 1 <= value <= 30
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionAcuerdo"
        )

    def coaseguradora(
        self, file_idx: int, exception_sheet: str, exception_col: str
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        exception_col: pd.Series = exception_df[exception_col].dropna()
        file_col: pd.Series = data_frame.iloc[:, file_idx]
        data_frame["is_valid"] = (file_col.isin(exception_col)) | (pd.isna(file_col))

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, file_idx, "CompañiaCoaseguradora"
        )

    def only_two_options(self, col_idx: int, options: list[str], new_sheet: str) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["id_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: (value in options) or (pd.isna(value))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["id_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def no_white_spaces(self, col_idx: int, new_sheet: str) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(
                lambda value: (pd.isna(value)) or not (bool(re.search(r"\s\s+", value)))
            )
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def percentage_format(self, col_idx: int, can_be_null: bool) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)

        def validate_format(value: str) -> bool:
            value = value.replace(" ", "")
            normal_percentage = bool(re.search(r"^\d+\.\d{1,2}$", value))
            concat_percentage = bool(re.search(r"^\d{2}%;\d{2}%$", value))
            if not can_be_null:
                return normal_percentage or concat_percentage or value == "1"
            else:
                is_nan: bool = value.lower() == "nan"
                return normal_percentage or concat_percentage or is_nan

        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(lambda value: validate_format(value))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "FormatoPorcentaje"
        )

    def identification_pagos_iaxis(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, "OTRAS EXCEPCIONES"
        )
        exception_list: pd.Series = (
            exception_df.iloc[:, 5].dropna().astype(str).to_list()
        )

        ## Create a  subfunction to validate the identification
        def validate_identification(identificador_pagos: str, radicado: str) -> bool:
            if bool(re.search(r"(^\s+|\s+$|\s{2,})", identificador_pagos)):
                return False
            else:
                return (identificador_pagos != "nan") or (radicado in exception_list)

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_identification(str(row.iloc[75]), str(row.iloc[2])),
            axis=1,
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, [12, 75], "IdentificacionPagosIaxis"
        )

    def need_exception(
        self,
        col_idx: int,
        exception_sheet: str,
        exception_idx: int,
        new_sheet: str,
        list_sheet: str,
        list_idx: int,
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        ## List data frame
        list_df: pd.DataFrame = self.read_excel(self.exception_file, list_sheet)
        list_col: list[str] = list_df.iloc[:, list_idx].dropna().astype(str).to_list()
        ## Exception values
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        exception_col: list[str] = (
            exception_df.iloc[:, exception_idx].dropna().astype(str).to_list()
        )
        file_col: pd.Series = data_frame.iloc[:, col_idx].astype(str)
        data_frame["is_valid"] = (file_col.isin(exception_col)) | (
            file_col.isin(list_col)
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def banks_validation(self) -> str:
        ## Data frames
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        list_df: pd.DataFrame = self.read_excel(self.exception_file, "LISTAS")
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, "OTRAS EXCEPCIONES"
        )
        new_list_df: pd.DataFrame = list_df.iloc[:, 1:3].dropna()
        exception_list: list[str] = (
            exception_df.iloc[:, 1].dropna().astype(str).to_list()
        )
        col_1_name: str = data_frame.columns[64]
        col_2_name: str = new_list_df.columns[0]
        merged_df: pd.DataFrame = data_frame.merge(
            new_list_df,
            how="left",
            left_on=col_1_name,
            right_on=col_2_name,
            suffixes=("_PAGOS", "_LISTAS"),
        )
        merged_df["is_valid"] = (merged_df.iloc[:, 65] == merged_df.iloc[:, -1]) | (
            merged_df.iloc[:, 64].astype(str).isin(exception_list)
        )
        inconsistencies: pd.DataFrame = merged_df[~merged_df["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, 64, "ValidacionBancos")

    def mandatory_desempleo(self, new_sheet: str, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)

        ## Sub function to validate
        def validation(desempleo: str, character: str) -> bool:
            tomadores_allowed: list[str] = ["FONDO NACIONAL DEL AHORRO"]
            is_desempleo = desempleo in tomadores_allowed
            is_valid_character = character == "SI" or character == "NO"
            first_validation = is_desempleo and is_valid_character
            second_validation = not is_desempleo and character == "nan"
            return first_validation or second_validation

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validation(
                str(row.iloc[15]),  # Tomador column
                str(row.iloc[col_idx]),  # Special column
            ),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, [12, col_idx], new_sheet)

    def no_empty(self, col_idx: int, option: str, new_sheet: str) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)

        ## Sub function to validate

        def validate_empty(value: str) -> bool:
            return value == "nan" or value == option

        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: validate_empty(str(value))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, [col_idx], new_sheet)

    def check_sarlaf(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)

        ## Sub function to validate
        def validate_sarlaf(sarlaf: str, bien_diligenciado: str) -> bool:
            if sarlaf == "SI":
                return bien_diligenciado == "X"
            else:
                return bien_diligenciado == "nan"

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_sarlaf(
                str(row.iloc[85]), str(row.iloc[86])  # Special column
            ),
            axis=1,
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, [85, 86], "CheckBeneficiarioSarlaf"
        )

    def fecha_vencimiento(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, "OTRAS EXCEPCIONES"
        )
        exception_list: list[str] = (
            exception_df.iloc[:, 3].dropna().astype(str).to_list()
        )

        ## Sub function to validate the expiration date
        def validate_date(ramo: str, expiration_date: str) -> bool:
            if ramo == "DESEMPLEO":
                return bool(re.search(r"^\d{2}/\d{2}/\d{4};\d{1,12}$", expiration_date))
            else:
                return (expiration_date == "nan") or ramo in exception_list

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_date(
                str(row.iloc[12]),  ## Ramo
                str(row.iloc[97]),  ## Fecha vencimiento
            ),  # Ramo
            axis=1,
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, [12, 97], "FechaVencimiento"
        )

    def evento_cinco(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame["EVENTO 5"].apply(
            lambda value: (pd.isna(value)) | (value == "SI" or value == "NO")
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, 110, "ValidacionEventoCinco"
        )

    def sap(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(self.exception_file, "LISTAS")
        exception_list: list[str] = exception_df["SAP"].dropna().astype(str).to_list()

        ## Sub function for making the validation
        def validate_number(sap: str) -> bool:
            try:
                int(sap)
                return True
            except ValueError:
                return sap in exception_list

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_number(str(row.iloc[77])),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, 77, "ValidacionSap")

    def otros_documentos(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        polizas: list[str] = ["3400004306", "3400003706", "3400003704", "3400004407"]
        allowed: list[str] = ["SI", "NO", "NA"]
        data_frame = data_frame[data_frame.iloc[:, 11].astype(str) == "334"]

        ## Sub function to validate the cell format
        def validate_cell_format(poliza: str, value: str) -> bool:
            if poliza in polizas:
                return value in allowed
            else:
                return value == "nan"

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_cell_format(
                str(row.iloc[6]),  # Poliza
                str(row.iloc[103]),  # Otros documentos
            ),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, [6, 103], "ValidacionOtrosDocumentos"
        )

    def concepto(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(self.exception_file, "LISTAS")
        exception_list: list[str] = (
            exception_df["CONCEPTO"].dropna().astype(str).to_list()
        )

        ## Sub function to validate the concepto
        def validate_concepto(concepto: str) -> bool:
            return concepto in exception_list

        data_frame["is_valid"] = data_frame["CONCEPTO"].apply(
            lambda value: validate_concepto(str(value))
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, 35, "ValidacionConcepto")


## Set global variables
validation_group: Optional[FirstValidationGroup] = None


def main(params: dict) -> bool:
    try:
        global validation_group

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        exception_file: str = params.get("exception_file")

        ## Pass the values to the constructor in the main class
        validation_group = FirstValidationGroup(
            file_path, sheet_name, inconsistencies_file, exception_file
        )
        return True
    except Exception as e:
        return f"ERROR: {e}"


def validate_empty_cols(incomes: dict) -> str:
    try:
        col_idx = int(incomes.get("col_idx"))
        is_mandatory = incomes.get("is_mandatory")

        validate: str = validation_group.validate_empty_col(col_idx, is_mandatory)
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_number_type(params: dict) -> str:
    try:
        ## Set local variables
        index = int(params.get("col_idx"))

        validate: str = validation_group.number_type(index)
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_date_type(incomes: dict) -> str:
    try:
        ## Set local variables
        index = int(incomes.get("col_idx"))

        validate: str = validation_group.date_type(index)
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_length(incomes: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(incomes.get("col_idx"))
        length = int(incomes.get("length"))

        validation: str = validation_group.value_length(col_idx, length)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_exception_list(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        exception_col_name = params.get("exception_col_name")
        exception_sheet = params.get("exception_sheet")
        new_sheet = params.get("new_sheet")

        validate: str = validation_group.validate_exception_list(
            col_idx, exception_col_name, exception_sheet, new_sheet
        )
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_special_characters(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.no_special_characters(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_month(params: dict) -> str:
    try:
        ## Set local variables
        date_idx = int(params.get("date_idx"))
        month_idx = int(params.get("month_idx"))
        exception_sheet = params.get("exception_sheet")
        exception_idx = int(params.get("exception_idx"))

        validation: str = validation_group.month_depends_on_date(
            date_idx, month_idx, exception_sheet, exception_idx
        )
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_numero_radicado(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.radicado_format(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_acuerdo_range(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.acuerdo_range(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_compania_coaseguradora(params: dict) -> str:
    try:
        ## Set local variables
        file_idx = int(params.get("file_idx"))
        exception_sheet = params.get("exception_sheet")
        exception_col = params.get("exception_col")

        validation: str = validation_group.coaseguradora(
            file_idx, exception_sheet, exception_col
        )
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_only_two_options(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        options = params.get("options")
        new_sheet = params.get("new_sheet")
        validation: str = validation_group.only_two_options(col_idx, options, new_sheet)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_no_white_spaces(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        new_sheet = params.get("new_sheet")

        validation: str = validation_group.no_white_spaces(col_idx, new_sheet)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_percentage_format(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        can_be_null = bool(params.get("can_be_null"))

        validation: str = validation_group.percentage_format(col_idx, can_be_null)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_identification_pagos_iaxis() -> str:
    try:
        validation: str = validation_group.identification_pagos_iaxis()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_need_exception(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        exception_sheet = params.get("exception_sheet")
        exception_idx = int(params.get("exception_idx"))
        new_sheet = params.get("new_sheet")
        list_sheet = params.get("list_sheet")
        list_idx = int(params.get("list_idx"))

        validation: str = validation_group.need_exception(
            col_idx, exception_sheet, exception_idx, new_sheet, list_sheet, list_idx
        )
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_banks() -> str:
    try:
        validation: str = validation_group.banks_validation()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_mandatory_desempleo(params: dict) -> str:
    try:
        ## Set local variables
        new_sheet = params.get("new_sheet")
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.mandatory_desempleo(new_sheet, col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_not_empty(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        option = params.get("option")
        new_sheet = params.get("new_sheet")
        validation: str = validation_group.no_empty(col_idx, option, new_sheet)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_check_sarlaf() -> str:
    try:
        validation: str = validation_group.check_sarlaf()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_fecha_vencimiento() -> str:
    try:
        validation: str = validation_group.fecha_vencimiento()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_evento_5() -> str:
    try:
        validation: str = validation_group.evento_cinco()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_sap() -> str:
    try:
        validation: str = validation_group.sap()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_otros_documentos() -> str:
    try:
        validation: str = validation_group.otros_documentos()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_concepto() -> str:
    try:
        validation: str = validation_group.concepto()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBasePagos.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE PAGOS.xlsx",
    }
    print(main(params))
    print(validate_sap())
