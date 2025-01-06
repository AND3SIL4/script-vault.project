import pandas as pd  # type: ignore
from typing import Optional


class MeshValidation:
    def __init__(
        self,
        file_path: str,
        sheet_name: str,
        exception_file: str,
        inconsistencies_file: str,
    ):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.exception_file = exception_file
        self.inconsistencies_file = inconsistencies_file

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
            inconsistencies, [col_idx], "ValidacionValorTipoNúmero"
        )

    except ValueError:
        return False


def validate_date_type(col_idx: str) -> str:
    """Method to validate if a column index is a date"""
    try:
        col_idx = int(col_idx)
        data_frame: pd.DataFrame = mesh_validation.read_excel(
            mesh_validation.file_path,
            mesh_validation.sheet_name,
        )
        data_frame["is_date"] = pd.to_datetime(
            data_frame.iloc[:, col_idx], errors="coerce"
        )
        # Validate if there is inconsistencies
        inconsistencies = data_frame[~data_frame["is_date"]]
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [col_idx], "ValidacionValorTipoFecha"
        )
    except Exception as e:
        return (False, f"Error: {e}")


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
            lambda poliza: str(poliza).startswith("31") or str(poliza).startswith("35")
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
        print(inconsistencies)
        return mesh_validation.validate_inconsistencies(
            inconsistencies, [col_idx], "ValidacionEspacios"
        )
    except Exception as e:
        return (False, f"Error: {e}")


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\PROPUESTA DE PAGO (23-10-2024).xlsx",
        "sheet_name": "Propuesta",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\EXCEPCIONES BASE PAGOS RED ASISTENCIAL.xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\InconsistenciasBasePagosRedAsistencial.xlsx",
    }
    # Instance the main variables for the main class
    main_instance = main(params)
    print(main_instance)

    # Instance an alone function with its params to test it
    print(validate_spaces("2"))
