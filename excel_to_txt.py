import pandas as pd
from openpyxl.utils import column_index_from_string


def excel_to_txt(package_numbers: list, row_numbers: list,
                 output_path: str, excel_path: str) -> None:
    """
    Converts specific rows and columns from an Excel file to a text file
    based on package numbers and row indices.

    Args:
        package_numbers (list): A list of package numbers (values in the
            "Paczka" column) to filter rows by.
        row_numbers (list): A list of row indices (zero-based) to be included
            in the output.
        output_path (str): The path where the output text file will be saved.
        excel_path (str): The path to the Excel file containing the data
            to be processed.

    Returns:
        None: This function writes the filtered data directly to a text file
            and does not return any value.
    """
    df = pd.read_excel(excel_path)

    df_filtered = df[
        df.index.isin(row_numbers) & df['Paczka'].isin(package_numbers)]

    with open(output_path, 'w', encoding='utf-8') as f:
        for idx, (_, row) in enumerate(df_filtered.iterrows(), 1):
            f.write(f"Liczba pozycyjna: {idx}\n")
            f.write(f"Treść zgłoszenia: {row.iloc[
                column_index_from_string('E') - 1]}\n")
            f.write(f"Moduł: {row.iloc[column_index_from_string('D') - 1]}\n")
            f.write(f"Streszczenie korespondencji: {row.iloc[
                column_index_from_string('F') - 1]}")
            f.write("-"*40 + "\n\n")


package_numbers = [2]
row_numbers = [10, 11]
output_path = 'Próbka zgłoszenia ver.2.txt'
excel_path = 'Próbka zgłoszenia ver.2.xlsx'

excel_to_txt(package_numbers, row_numbers, output_path, excel_path)
