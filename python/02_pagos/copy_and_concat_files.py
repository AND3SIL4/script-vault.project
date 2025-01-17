import pandas as pd  # type: ignore
import os


def main(params: dict):
    try:
        ##Set initial variables and values
        otros_ramos_file = params.get("otros_ramos_file")
        desempleo_file = params.get("desempleo_file")
        sheet_otros_ramos = params.get("sheet_otros_ramos")
        sheet_desempleo = params.get("sheet_desempleo")
        begin_date = params.get("begin_date")
        cut_off_date = params.get("cut_off_date")
        col_idx = int(params.get("col_idx"))
        destination_path = params.get("destination_path")

        ##Validate if all inputs required are present
        if not all(
            [
                otros_ramos_file,
                desempleo_file,
                sheet_otros_ramos,
                sheet_desempleo,
                begin_date,
                cut_off_date,
                col_idx,
            ]
        ):
            return "ERROR: an input required param is missing"

        ##Make a date type to filter the files
        begin_date = pd.to_datetime(begin_date, format="%d/%m/%Y")
        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        ##Read the books and make a filter
        otros_ramos_df: pd.DataFrame = pd.read_excel(
            otros_ramos_file, sheet_name=sheet_otros_ramos, engine="openpyxl"
        )
        desempleo_df: pd.DataFrame = pd.read_excel(
            desempleo_file, sheet_name=sheet_desempleo, engine="openpyxl"
        )
        otros_gastos: pd.DataFrame = pd.read_excel(
            otros_ramos_file, sheet_name="OGDS", engine="openpyxl"
        )

        otros_ramos_df: pd.DataFrame = otros_ramos_df.dropna(
            subset=[otros_ramos_df.columns[0]]
        ).iloc[:, :111]
        desempleo_df: pd.DataFrame = desempleo_df.dropna(
            subset=[desempleo_df.columns[1]]
        ).iloc[:, :111]
        otros_gastos: pd.DataFrame = otros_gastos.dropna(
            subset=[otros_gastos.columns[0]]
        ).iloc[:, :111]

        ## Unique cols name
        desempleo_df.columns = otros_ramos_df.columns
        otros_gastos.columns = otros_ramos_df.columns
        ##Link the files previously filtered
        base_pagos = pd.concat(
            [desempleo_df, otros_ramos_df, otros_gastos], ignore_index=True
        )

        ##Convert the column to date type
        base_pagos.iloc[:, col_idx] = pd.to_datetime(
            base_pagos.iloc[:, col_idx], format="%d/%m/%Y"
        )

        ##Make a filter
        base_pagos = base_pagos[
            (base_pagos.iloc[:, col_idx] >= begin_date)
            & (base_pagos.iloc[:, col_idx] <= cut_off_date)
        ]

        ##Save changes into a temp folder
        base_pagos.to_excel(destination_path, index=False, sheet_name="PAGOS")
        return "Temp file created successfully"

    except Exception as e:
        return f"Error: {e}"


if __name__ == "__main__":
    params = {
        "otros_ramos_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BDD PAGOS RECONOCIMIENTO POLIZAS DE VIDA PAGOS 2024 - OTROS RAMOS.xlsx",
        "desempleo_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BDD PAGOS RECONOCIMIENTO POLIZAS DE VIDA – PAGOS 2024 - DESEMPLEO.xlsx",
        "sheet_otros_ramos": "2024",
        "sheet_desempleo": "2024 DESEMPLEO",
        "begin_date": "01/01/2024",
        "cut_off_date": "28/10/2024",
        "col_idx": 72,
        "destination_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
    }
    print(main(params))
