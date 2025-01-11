import pandas as pd  # type: ignore
from typing import Tuple


def get_initial_date(incomes: dict) -> Tuple[bool, str]:
    try:
        # Set local variables
        propuesta_file: str = incomes.get("propuesta_file")
        propuesta_sheet: str = incomes.get("propuesta_sheet")
        col_idx = int(incomes.get("col_idx"))

        # Read the Propuesta de Pago file
        data_frame: pd.DataFrame = pd.read_excel(
            propuesta_file,
            propuesta_sheet,
            engine="openpyxl",
        )
        # Get the initial date from the "Radicado casa matriz" column
        radicado_list: list[str] = (
            data_frame.iloc[:, col_idx].dropna().astype(str).str[:4].to_list()
        )

        # Convert string list into int list
        radicado_list = [int(radicado) for radicado in radicado_list]
        # Get the minimum radicado
        initial_date = str(min(radicado_list))
        return True, initial_date
    except Exception as e:
        return False, str(e)


print(
    get_initial_date(
        {
            "propuesta_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\PROPUESTA DE PAGO 1 Y 2  (02-01-2025).xlsx",
            "propuesta_sheet": "Propuesta",
            "col_idx": "2",
        }
    )
)
