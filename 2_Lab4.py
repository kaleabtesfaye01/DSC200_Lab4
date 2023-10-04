"""
Name: Kaleab Alemu and Manogya Aryal
DSC 200
Lab 4: Working with Excel Files

This program reads data from "Lab4Data.xlsx" which contains data collected from countries regarding child abuse in those
respective countries and creates a CSV file containing the country name, category of child abuse and category total

Due Date: Oct 4, 2023
"""

import openpyxl as op
import csv


class DSCLab5:
    def __init__(self,csvFile):
        self.csvFile = csvFile

    def get_categories(ws):
        category_names = []
        for cols in ws.iter_rows(min_row=5, max_row=7, min_col=5, max_col=31):
            row = []
            for cell in cols:
                currVal = cell.value
                for merged_cells in ws.merged_cells.ranges:
                    if cell.coordinate in merged_cells:
                        currVal = merged_cells.start_cell.value
                if cell.row == 6 and currVal in category_names[0]:
                    currVal = ''
                row.append(currVal)
            category_names.append(row)

        categories = []
        for i in range(len(category_names[0])):
            category = ''
            for j in range(len(category_names)):
                category_names[j][i] = category_names[j][i].replace('\n', '')
                category = category + '_' + category_names[j][i] if category != '' else category_names[j][i]
            categories.append(category)

        return list(dict.fromkeys(categories))


    # This function "get_values", given the worksheet "Table 9" as a parameter, returns a list of lists containing all the
    # values for each category of child abuse for each country
    def get_values(ws):
        values = list()  # initialize our values list to an empty list

        for row in ws['E15:AF211']:  # iterate through cells E15:AF211 to extract all the relevant values
            value = []  # holds the list of values for one country
            for cell in row:  # iterate through cells in each row
                # if the value is an integer or a floating point number or an en dash, we add it to our list of values
                if type(cell.value) == int or type(cell.value) == float or cell.value == chr(0x2013):
                    value.append(cell.value)
            values.append(value)  # append the values in one row into the values list
        return values


    # This function "get_countries", given the worksheet "Table 9" as a parameter, returns the list of countries in the
    # worksheet
    def get_countries(ws):
        countries = list()  # initialize our countries list to an empty list

        for row in ws['B15:B211']:  # iterate through cells B15:B211 to extract the country names
            for cell in row:  # iterate though the cells in each row
                countries.append(cell.value)  # add the value of each cell into our countries list
        return countries


    def write_csv(self,countries, categories, values):
        csv_file = self.csvFile
        heading = ['CountryName', 'CategoryName', 'CategoryTotal']

        with open(csv_file, 'w', newline ='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(heading)

            # iterate through the list to get the rows
            for country in countries:
                for category, value_list in zip(categories, values):
                    for value in value_list:
                        csvwriter.writerow([country, category, value])


def main():
    # load our workbook
    wb = op.load_workbook('./data/Lab4Data.xlsx')

    # open the active worksheet, "Table 9"
    ws = wb.active
    lab5 = DSCLab5('aryalm1_alemuk1_lab4.csv')
    # call the get_countries, get_values and get_categories functions and enter the results as parameters for our
    # write_csv function
    # write_csv(get_countries(ws), get_values(ws), get_categories(ws))
    countries = lab5.get_countries(ws)
    values = lab5.get_values(ws)
    categories = lab5.get_categories(ws)

    lab5.write_csv(countries, categories, values)

main()  # run our script
