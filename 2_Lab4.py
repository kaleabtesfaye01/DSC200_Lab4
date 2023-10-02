import openpyxl as op
import csv


def get_categories(ws):
    pass


def get_values(ws):
    values = []

    for row in ws['E15:AF223']:
        value = []
        for cell in row:
            if type(cell.value) == int or type(cell.value) == float or cell.value == chr(0x2013):
                value.append(cell.value)
        values.append(value)
    return values


def get_countries(ws):
    pass


def write_csv(countries, categories, values):
    pass


def main():
    wb = op.load_workbook('./data/Lab4Data.xlsx', read_only=True, data_only=True)

    ws = wb.active

    get_values(ws)

main()