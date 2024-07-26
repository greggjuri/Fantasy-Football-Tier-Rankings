import pandas as pd


def csv_to_excel(pos):
    read_file = pd.read_csv(
        f"D:/Documents/Fantasy Football/2024/UDK - Position Rankings - {pos}.csv"
    )
    read_file.to_excel(
        f"D:/Documents/Fantasy Football/2024/UDK - Position Rankings - {pos}.xlsx",
        index=None,
        header=True,
    )
