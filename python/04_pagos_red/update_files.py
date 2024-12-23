def update_validador_pagos_file(data_frame: pd.DataFrame) -> None:
    """Method to update the propuesta de pagos file with the new data frame"""
    try:
        # 1. Get only the needed columns from the data frame
        data_frame.iloc[:, 1:3]
        return data_frame
        # 2. Get the current year
        year = datetime.now().year
        book = load_workbook(values_validation.historic_file)
        sheet_names = book.sheetnames

        if str(year) in sheet_names:
            print(f"The sheet: {year} already exist in book")

        with pd.ExcelWriter(
            values_validation.historic_file,
            engine="openpyxl",
            if_sheet_exists="replace",
            mode="a",
        ) as writer:
            writer.book = book
            writer.sheets = {
                sheet.title: sheet for sheet in book.worksheets
            }  # Mapea las hojas existentes
            data_frame.to_excel(
                values_validation.historic_file, sheet_name=str(year), index=False
            )

    except Exception as e:
        return False, f"Error: {e}"
    

def final_file():
    final_df: pd.DataFrame = pd.concat(
            [historical_df, filled_df], ignore_index=True
        )
        # Save the final file into temp file folder
        final_df.to_excel(
            values_validation.temp_file,
            sheet_name=values_validation.sheet_name,
            index=False,
        )