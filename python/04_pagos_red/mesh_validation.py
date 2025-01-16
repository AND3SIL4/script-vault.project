import pandas as pd  # type: ignore
from typing import Optional, Tuple


class MeshValidation:
    def __init__(
        self,
        file_path: str,
        sheet_name: str,
        exception_file: str,
        inconsistencies_file: str,
        acm_report: str,
    ):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.exception_file = exception_file
        self.inconsistencies_file = inconsistencies_file
        self.acm_report = acm_report

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        data_frame: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        # Get the data frame using the name of the column with the index 2
        # To avoid the NaN into data frame with the aim the next validations
        df = data_frame.dropna(subset=[data_frame.columns[2]])
        return df

    def save_inconsistencies_file(self, df: pd.DataFrame, new_sheet: str) -> bool:
        """Method to save the inconsistencies in a new sheet or update an existing one"""
        with pd.ExcelFile(self.inconsistencies_file, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, engine="openpyxl", sheet_name=new_sheet)
                df = pd.concat([existing, df], ignore_index=True)

        with pd.ExcelWriter(
            self.inconsistencies_file,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            df.to_excel(writer, sheet_name=new_sheet, index=False)
            return True

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

    def transform_acm_report(self, acm_report: pd.DataFrame) -> pd.DataFrame:
        """Method to transform the ACM report"""
        acm_report: pd.DataFrame = acm_report.iloc[3:, 1:]
        acm_report.columns = acm_report.iloc[0]
        acm_report = acm_report.iloc[1:].reset_index(drop=True)
        # Select only the columns needed
        acm_report = acm_report[["id cuenta", "prefijo factura", "factura"]]
        return acm_report

    def validate_vs_coaseguro(
        self, data_frame: pd.DataFrame, column1: int, column2: int, sheet_name: str
    ) -> Tuple[bool, str]:
        """Method to validate the tipo expedición poliza"""
        # Validate if the dataframe is empty
        if data_frame.empty:
            return False

        try:
            # If the validation made is a number type then convert to integers and the string to validate
            if column1 == 16:
                data_frame.iloc[:, column1] = (
                    data_frame.iloc[:, column1].astype(int).astype(str)
                )
            # Validate if the tipo expedición poliza is correct
            data_frame["is_valid"] = data_frame.apply(
                lambda row: str(row.iloc[column1]) == str(row.iloc[column2]), axis=1
            )
            # Validate inconsistencies
            inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]].copy()
            report_inconsistencies = self.validate_inconsistencies(
                inconsistencies, [column1, column2], sheet_name
            )
            print(report_inconsistencies)
            return (True, "Validacion realizada correctamente")
        except Exception as e:
            return (False, str(e))

    def validate_sum_percentage(
        self, data_frame: pd.DataFrame, positiva: int, coaseguradora: int
    ) -> Tuple[bool, str]:
        """Method to validate the sum of the percentages"""
        try:
            data_frame["percentage_valid"] = data_frame.apply(
                lambda row: row.iloc[positiva] + row.iloc[coaseguradora] == 1, axis=1
            )
            # Validate inconsistencies
            inconsistencies: pd.DataFrame = data_frame[
                ~data_frame["percentage_valid"]
            ].copy()
            report_inconsistencies = self.validate_inconsistencies(
                inconsistencies, [positiva, coaseguradora], "ValidacionPorcentajes"
            )
            print(report_inconsistencies)
            return (True, "Validacion realizada correctamente")
        except Exception as e:
            return (False, str(e))

    def validate_positiva_plus_movimiento(
        self, data_frame: pd.DataFrame, vr_movimiento: int, positiva_percentage: int
    ) -> Tuple[bool, str]:
        """Method to validate the sum of the VR movimiento and the positiva percentage"""
        try:
            data_frame["is_valid"] = data_frame.apply(
                lambda row: row.iloc[vr_movimiento] * row.iloc[positiva_percentage] > 0,
                axis=1,
            )
            # Validate inconsistencies
            inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]].copy()
            report_inconsistencies = self.validate_inconsistencies(
                inconsistencies,
                [vr_movimiento, positiva_percentage],
                "ValidacionVRMovimiento",
            )
            print(report_inconsistencies)
            return (True, "Validacion realizada correctamente")
        except Exception as e:
            return (False, str(e))


mesh_validation: Optional[MeshValidation] = None


def main(params: dict) -> tuple:
    """Main method to validate and save inconsistencies"""
    global mesh_validation
    try:
        if mesh_validation is None:
            mesh_validation = MeshValidation(
                file_path=params.get("file_path"),
                sheet_name=params.get("sheet_name"),
                exception_file=params.get("exception_file"),
                inconsistencies_file=params.get("inconsistencies_file"),
                acm_report=params.get("acm_report"),
            )
        return True, f"Atributos de clase '{main.__name__}' inicializados correctamente"
    except Exception as e:
        return (False, f"Error: {e}")


def validate_is_number(col_idx: str) -> str:
    """Method to validate if a column index is a number"""
    try:
        col_idx = int(col_idx)

        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )

        data_frame["is_number"] = data_frame.iloc[:, col_idx].apply(
            lambda value: str(value).replace(".", "").replace(",", "").isdigit()
        )

        # Validate if there is inconsistencies
        inconsistencies = data_frame[~data_frame["is_number"]]
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionValorTipoNúmero"
        )

    except ValueError:
        return False


def validate_date_type(col_idx: str) -> str:
    """Method to validate if a column index contains date values."""
    try:
        # Convert column index to integer
        col_idx = int(col_idx)

        # Read the Excel file into a DataFrame
        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )

        # Attempt to convert the column to datetime
        data_frame["is_date"] = pd.to_datetime(
            data_frame.iloc[:, col_idx], errors="coerce"
        ).notna()

        # Find inconsistencies (non-date values)
        inconsistencies = data_frame[~data_frame["is_date"]]

        # Validate and return the result
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionValorTipoFecha"
        )
    except Exception as e:
        return f"Error: {e}"


def validate_siniestro_number() -> str:
    try:
        # Set local variables
        col_idx: int = 0  # No. SINIESTRO

        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )

        # Subfunction to validate multiple columns that make up the siniestro column
        def validate_siniestro(siniestro: str, documento: str, fecha: pd.Timestamp):
            """
            This function makes a validation to enure that the siniestro number is valid
            The siniestro number is make up with two columns specify bellow
                siniestro: No. SINIESTRO = documento + fecha (0)
                documento: DOCUMENTO RIESGO (Asegurado) (18)
                fecha: FECHA SINIESTRO (27)
            """
            # Convert the date passed into string with the format: "DDMMYYYY"
            fecha = fecha.strftime("%d%m%Y")
            # Check if the siniestro number is in the exception list
            # Validate if all the values are present and passed
            if siniestro or documento or fecha:
                # Check if the siniestro number is composed by two columns
                # Create a key composted by DOCUMENTO RIESGO + FECHA SINIESTRO
                siniestro_composed: str = documento + fecha
                # Check if the siniestro and siniestro composed are the same and return it
                return siniestro == siniestro_composed
            else:
                return False

        data_frame["is_siniestro_valid"] = data_frame.apply(
            lambda row: validate_siniestro(
                str(row.iloc[0]),  # No. SINIESTRO
                str(int(row.iloc[18])),  # DOCUMENTO RIESGO (Asegurado)
                row.iloc[27],  # FECHA SINIESTRO
            ),
            axis=1,
        )

        # Validate if there is inconsistencies
        inconsistencies = data_frame[~data_frame["is_siniestro_valid"]].copy()
        # Get the exception list
        exception_df: pd.DataFrame = pd.read_excel(
            mesh_validation.exception_file,
            sheet_name="EXCEPCIONES GENERALES",
            engine="openpyxl",
            dtype=str,
        )
        exception_list: list[str] = (
            exception_df.iloc[:, 0].dropna().astype(str).to_list()
        )
        inconsistencies["is_exception"] = inconsistencies.iloc[:, 0].isin(
            exception_list
        )
        inconsistencies = inconsistencies[~inconsistencies["is_exception"]]
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [col_idx], "ValidacionNumeroSiniestro"
        )
    except Exception as e:
        return (False, f"Error: {e}")


def validate_poliza_number() -> str:
    try:
        # Set local variables
        col_idx: int = 6  # No. POLIZA
        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )
        # Validate if there is inconsistencies
        data_frame["is_poliza_number_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda poliza: str(poliza).startswith("331")
            or str(poliza).startswith("335")
        )
        inconsistencies = data_frame[~data_frame["is_poliza_number_valid"]].copy()
        # Create a exception data frame
        exception_df: pd.DataFrame = pd.read_excel(
            mesh_validation.exception_file,
            sheet_name="EXCEPCIONES GENERALES",
            engine="openpyxl",
            dtype=str,
        )
        # Get the exception list from the exception data frame
        exception_list: list[str] = (
            exception_df.iloc[:, 1].dropna().astype(str).to_list()
        )
        # Add the exception list to the inconsistencies data frame
        inconsistencies["exception_poliza"] = (
            inconsistencies.iloc[:, col_idx]
            .astype(int)
            .astype(str)
            .isin(exception_list)
        )
        # Remove the exception polizas from the inconsistencies data frame
        inconsistencies = inconsistencies[~inconsistencies["exception_poliza"]]
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionNumeroPoliza"
        )
    except Exception as e:
        return (False, f"Error: {e}")


def validate_spaces(col_idx: str) -> str:
    """Method to validate if a column index has spaces"""
    try:
        col_idx = int(col_idx)
        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )
        # Validate spaces using regex expression: s\
        # Validate double withe spaces using regex expression: s\{2,}
        data_frame["has_spaces"] = (
            data_frame.iloc[:, col_idx].astype(str).str.contains(r"\s", regex=True)
            == False
        )
        # Validate if there is inconsistencies
        inconsistencies = data_frame[~data_frame["has_spaces"]].copy()
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [col_idx], "ValidacionEspacios"
        )
    except Exception as e:
        return (False, f"Error: {e}")


def validate_observaciones_col() -> str:
    try:
        # Get the PROPUESTA PAGOS data frame
        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )
        # Get the ACM report data frame
        acm_df: pd.DataFrame = pd.read_excel(
            mesh_validation.acm_report,
            sheet_name="FCT_RS_REPORTE_WS_AUDITORIA",
            engine="openpyxl",
            dtype=str,
        )
        # Transform the data frame
        df_transformed: pd.DataFrame = mesh_validation.transform_acm_report(acm_df)

        # Get columns to make the merge
        left_col = data_frame.columns[2]
        right_col = df_transformed.columns[0]

        # Set the columns type
        data_frame[left_col] = data_frame[left_col].astype(int).astype(str)
        df_transformed[right_col] = df_transformed[right_col].astype(str)

        # Merge the data frames
        merged_df: pd.DataFrame = pd.merge(
            data_frame,
            df_transformed,
            how="left",
            left_on=left_col,
            right_on=right_col,
            suffixes=("_propuesta_pago", "_acm_report"),
        )

        # Subfunction to validate the key composed
        def validate_key(
            observaciones: str, prefijo_factura: str, id_factura: str
        ) -> bool:
            # Create a key composed
            key: str = prefijo_factura + id_factura
            # Validate if there is any NaN value and replace it
            if "nan" in key:
                key = key.replace("nan", "")

            # Check if the key is the same as the observations value is
            return key == observaciones

        merged_df["is_valid"] = merged_df.apply(
            lambda row: validate_key(
                str(row[merged_df.columns[62]]),  # OBSERVACIONES
                str(row[merged_df.columns[116]]),  # prefijo factura
                str(row[merged_df.columns[117]]),  # factura
            ),
            axis=1,
        )

        # Validate if there is inconsistencies
        inconsistencies = merged_df[~merged_df["is_valid"]].copy()

        # Create a exception data frame
        exception_df: pd.DataFrame = pd.read_excel(
            mesh_validation.exception_file,
            sheet_name="EXCEPCIONES GENERALES",
            engine="openpyxl",
            dtype=str,
        )
        # Get the exception list from the exception data frame
        exception_list: list[str] = (
            exception_df.iloc[:, 2].dropna().astype(str).to_list()
        )
        # Add the exception list to the inconsistencies data frame
        inconsistencies["is_exception"] = inconsistencies.iloc[:, 2].isin(
            exception_list
        )
        inconsistencies = inconsistencies[~inconsistencies["is_exception"]]
        print(inconsistencies)
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [62, 116, 117], "ValidacionObservacionesCol"
        )

    except Exception as e:
        return (False, f"Error: {e}")


def validate_coaseguro_sheet() -> Tuple[bool, str]:
    try:
        # Get main data frame
        propuesta_df: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )
        # Get COASEGURO data frame
        coaseguro_df: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.exception_file,
            "COASEGURO",
        )
        # Get only the important columns from coaseguro data frame
        # The first 6 columns with data
        coaseguro_df = coaseguro_df.iloc[:, :6]

        # Merge the data frames using the key: N° poliza
        # N° poliza propuesta pagos: index 6
        # N° poliza coaseguro sheet: index 0
        merged_df: pd.DataFrame = pd.merge(
            propuesta_df,
            coaseguro_df,
            how="left",
            left_on=propuesta_df.columns[6],
            right_on=coaseguro_df.columns[0],
            suffixes=("_propuesta", "_coaseguro"),
        )

        # * Validate "TIPO EXPEDICIÓN PÓLIZA"
        bool_tipo_exp, msg_tipo_exp = mesh_validation.validate_vs_coaseguro(
            merged_df, 42, 120, "ValidacionTipoExpedicionPoliza"
        )
        if not bool_tipo_exp:
            raise Exception(msg_tipo_exp)

        # * Validate "TOMADOR"
        bool_tomador, msg_tomador = mesh_validation.validate_vs_coaseguro(
            merged_df, 15, 116, "ValidacionTomadorCoaseguro"
        )
        if not bool_tomador:
            raise Exception(msg_tomador)

        # * Validate "TOMADOR"
        bool_doc_tomador, msg_doc_tomador = mesh_validation.validate_vs_coaseguro(
            merged_df, 16, 117, "ValidacionDocumentoTomador"
        )
        if not bool_doc_tomador:
            raise Exception(msg_doc_tomador)

        # * Validate sum of coaseguro POSITIVA PERCENTAGE AND COASEGURADORA PERCENTAGE
        bool_sum_coaseguro, msg_sum_coaseguro = mesh_validation.validate_sum_percentage(
            merged_df, 48, 50  # PORCENTAJE POSITIVA  # PORCENTAJE COASEGURADORA
        )
        if not bool_sum_coaseguro:
            raise Exception(msg_sum_coaseguro)

        # * Validate Valor Positiva MOVIMIENTO x %
        bool_val_positiva, msg_val_positiva = (
            mesh_validation.validate_positiva_plus_movimiento(
                merged_df, vr_movimiento=45, positiva_percentage=118
            )
        )
        if not bool_val_positiva:
            raise Exception(msg_val_positiva)

        return (True, "Validacion con hoja COASEGURO realizada correctamente")
    except Exception as e:
        return (False, f"Error: {e}")


def validate_empty(incomes: dict) -> Tuple[bool, str]:
    try:
        # Set local variables
        col = int(incomes.get("col"))
        is_empty: bool = incomes.get("is_empty")
        sheet_name: str = ""
        # Read the excel file
        df: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            sheet_name=mesh_validation.sheet_name,
        )
        if is_empty:
            # Check if the column must be empty
            df["is_valid"] = df.iloc[:, col].isna()
            sheet_name = "ValidacionColumnasVacias"
        else:
            # Check if the column must not be empty
            df["is_valid"] = ~df.iloc[:, col].isna()
            sheet_name = "ValidacionColumnasNoVacias"

        # Validate inconsistencies
        inconsistencies = df[~df["is_valid"]].copy()
        print(inconsistencies)
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [col], sheet_name=sheet_name
        )
    except Exception as e:
        return (False, f"Error: {e}")


def validate_using_list(incomes: dict) -> Tuple[bool, str]:
    try:
        # Set local variables
        col = int(incomes.get("col"))
        exception_sheet = incomes.get("exception_sheet")
        exception_col_name = incomes.get("exception_col_name")
        inconsistencies_sheet_name = incomes.get("inconsistencies_sheet_name")

        # Get main data frame
        df: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            sheet_name=mesh_validation.sheet_name,
        )
        # Get the list of the exception
        exception_df: list[str] = pd.read_excel(
            mesh_validation.exception_file,
            sheet_name=exception_sheet,
        )
        exception_list: list[str] = (
            exception_df[exception_col_name].dropna().astype(str).to_list()
        )
        # Check if the column must be in the exception list
        df["is_valid"] = df.iloc[:, col].isin(exception_list)
        # Validate inconsistencies
        inconsistencies = df[~df["is_valid"]].copy()
        print(inconsistencies)
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col, sheet_name=inconsistencies_sheet_name
        )

    except Exception as e:
        return (False, f"Error: {e}")


def validate_length(incomes: dict) -> Tuple[bool, str]:
    try:
        # Set local variables
        col = int(incomes.get("col"))
        length = int(incomes.get("length"))

        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            sheet_name=mesh_validation.sheet_name,
        )

        data_frame["is_valid"] = data_frame.iloc[:, col].apply(
            lambda x: len(str(int(x))) == length,
        )

        # Validate inconsistencies
        inconsistencies = data_frame[~data_frame["is_valid"]].copy()
        print(inconsistencies)
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col, sheet_name="ValidacionLongitudColumna"
        )
    except Exception as e:
        return (False, f"Error: {e}")


def validate_coaseguradora() -> str:
    try:
        # Column to validate COMPAÑIA COASEGURADORA
        col: int = 47
        percentage: int = 48

        # Get the data frame
        df: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            sheet_name=mesh_validation.sheet_name,
        )
        # Get the exception list
        exception_df: pd.DataFrame = pd.read_excel(
            mesh_validation.exception_file,
            sheet_name="OTRO",
        )
        exception_list: list[str] = (
            exception_df["COMPAÑIA COASEGURADORA"].dropna().astype(str).to_list()
        )

        def validate_coa(percentage_positiva: str, company: str):
            return float(percentage_positiva) == 1.0 or company in exception_list

        df["is_valid"] = df.apply(
            lambda row: validate_coa(
                percentage_positiva=str(row.iloc[percentage]),
                company=str(row.iloc[col]),
            ),
            axis=1,
        )
        # Validate inconsistencies
        inconsistencies = df[~df["is_valid"]].copy()
        # Return the inconsistencies
        return mesh_validation.validate_inconsistencies(
            inconsistencies, col, sheet_name="ValidacionCoaseguradora"
        )

    except Exception as e:
        return f"Error: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\PROPUESTA DE PAGO 1 Y 2  (02-01-2025).xlsx",
        "sheet_name": "Propuesta",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\EXCEPCIONES BASE PAGOS RED ASISTENCIAL.xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\InconsistenciasBasePagosRedAsistencial.xlsx",
        "acm_report": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\FCT_RS_REPORTE_WS_AUDITORIA.xlsx",
    }
    # Instance the main variables for the main class
    main_instance = main(params)
    print(main_instance)

    incomes = {
        "col": "47",
        "exception_sheet": "LISTAS",
        "exception_col_name": "COMPAÑIA COASEGURADORA",
        "inconsistencies_sheet_name": "VV",
    }
    # Instance an alone function with its params to test it
    print(validate_observaciones_col())
