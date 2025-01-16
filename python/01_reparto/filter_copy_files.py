import pandas as pd  # type: ignore
import os


def main(params):
    try:
        # Set variables
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        temp_file = params.get("temp_file")
        cut_off_date_input = params.get("cut_off_date")
        column_index = int(params.get("column_index"))
        start_date_input = params.get("start_date_input")

        # Validate parameters
        if not all(
            [
                file_path,
                sheet_name,
                temp_file,
                column_index,
                start_date_input,
                cut_off_date_input,
            ]
        ):
            return "ERROR: Missing one or more parameters."

        # Check if the file exists
        if not os.path.exists(file_path):
            return f"ERROR: File {file_path} not found."

        # Convert dates to datetime
        cut_off_date = pd.to_datetime(cut_off_date_input, format="%d/%m/%Y")
        start_date = pd.to_datetime(start_date_input, format="%d/%m/%Y")

        # Open file using pandas
        df = pd.read_excel(file_path, engine="openpyxl", sheet_name=sheet_name)
        otros_gastos = pd.read_excel(file_path, sheet_name="OGDS", engine="openpyxl")
        # Normalize columns
        df = df.iloc[:, :111]
        otros_gastos = otros_gastos.iloc[:, :111]

        ## Assign the columns of the first data frame
        otros_gastos.columns = df.columns
        df: pd.DataFrame = pd.concat([df, otros_gastos], ignore_index=True)

        # Convert the column to datetime using the column index
        df.iloc[:, column_index] = pd.to_datetime(
            df.iloc[:, column_index], format="%d/%m/%Y"
        )

        # Filter from start_date to cut_off_date
        filter_file = df[
            (df.iloc[:, column_index] >= start_date)
            & (df.iloc[:, column_index] <= cut_off_date)
        ]

        # Copy to temp file to make validations
        filter_file.to_excel(temp_file, index=False, sheet_name=sheet_name)
        return (True, "SUCCESS: file copied successfully")

    except Exception as e:
        return (False, f"ERROR: {e}")


if __name__ == "__main__":
    params = {
        # Set variables
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BASE DE REPARTO 2025.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "temp_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2025.xlsx",
        "start_date_input": "01/01/2025",
        "column_index": "24",
        "cut_off_date": "09/01/2025",
    }

    print(main(params))
