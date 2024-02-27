from datapackage import Package
import openpyxl
import csv
import os


def excel_to_csv(excel_file, sheet_name, csv_file):
    workbook = openpyxl.load_workbook(excel_file)

    sheet = workbook[sheet_name]

    with open(csv_file, 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)

        for row in sheet.iter_rows(values_only=True):
            csv_writer.writerow(row)


excel_file_path = 'arhive/6. ВРП на душу населения.xlsx'
sheet_tenge = 'тыс.тенге'
sheet_usd = 'в долларах США'
excel_to_csv(excel_file_path, sheet_tenge, 'data/tenge.csv')
excel_to_csv(excel_file_path, sheet_usd, 'data/usd0.csv')


def delete_last_column(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        for row in reader:
            writer.writerow(row[:-1])


delete_last_column('data/usd0.csv', 'data/usd.csv')


def rename_columns(input_file, output_file, new_column_names):
    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8',
                                                                 newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        header = next(reader)  # Read and store the original header

        # Check if the length of new_column_names matches the original header
        if len(new_column_names) != len(header):
            print("Number of column names does not match the number of columns in the file.")
            return

        # Rename the columns based on the provided names
        renamed_header = [new_column_names[i] if new_column_names[i] else header[i] for i in range(len(header))]

        writer.writerow(renamed_header)  # Write the header to the output file

        # Copy the remaining rows from the input file to the output file
        for row in reader:
            writer.writerow(row)


new_column_names = [
    'regions',
    'Quarter 2008', 'Half-Year 2008', '9 Months 2008', 'Year 2008',
    'Quarter 2009', 'Half-Year 2009', '9 Months 2009', 'Year 2009',
    'Quarter 2010', 'Half-Year 2010', '9 Months 2010', 'Year 2010',
    'Quarter 2011', 'Half-Year 2011', '9 Months 2011', 'Year 2011',
    'Quarter 2012', 'Half-Year 2012', '9 Months 2012', 'Year 2012',
    'Quarter 2013', 'Half-Year 2013', '9 Months 2013', 'Year 2013',
    'Quarter 2014', 'Half-Year 2014', '9 Months 2014', 'Year 2014',
    'Quarter 2015', 'Half-Year 2015', '9 Months 2015', 'Year 2015',
    'Quarter 2016', 'Half-Year 2016', '9 Months 2016', 'Year 2016',
    'Quarter 2017', 'Half-Year 2017', '9 Months 2017', 'Year 2017',
    'Quarter 2018', 'Half-Year 2018', '9 Months 2018', 'Year 2018',
    'Quarter 2019', 'Half-Year 2019', '9 Months 2019', 'Year 2019',
    'Quarter 2020', 'Half-Year 2020', '9 Months 2020', 'Year 2020',
    'Quarter 2021', 'Half-Year 2021', '9 Months 2021', 'Year 2021',
    'Quarter 2022', 'Half-Year 2022', '9 Months 2022', 'Year 2022',
    'Quarter 2023', 'Half-Year 2023', '9 Months 2023']


rename_columns('data/tenge.csv', 'data/tenge2.csv', new_column_names)
rename_columns('data/usd.csv', 'data/usd2.csv', new_column_names)


def remove_and_exclude(input_file, output_file):
    phrases_to_exclude = [
        'Валовой региональный продукт на душу населения1)',
        'тыс.тенге',
        'долл. США',
        '1) расчеты ВРП по кварталам начали прозводиться только с 2008 года.',
        '2) c 1990 по 2018 годы Туркестанская область и город Шымкент были в составе Южно-Казахстанской области.',
        "3) 13 февраля 2020 года обновлены данные отрасли сельское хозяйство по Павлодарской, Северо-Казахстанской и Туркестанской областям.",
        "4) Валовой региональный продукт за 1 полугодие 2022 года был пересчитан по Павлодарской, Северо-Казахстанской,  Туркестанской и Улытауской областям, в связи с пересчетом данных отрасли «Оптовая и розничная торговля; ремонт автомобилей и мотоциклов».",
        'Рассчитано по среднему курсу доллара США Национального банка РК'

    ]

    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        skip_line = False

        for row in reader:
            if any(phrase in ''.join(row) for phrase in phrases_to_exclude):
                skip_line = True
                continue
            if not skip_line:
                cleaned_row = [value for value in row if value.strip() != '']
                writer.writerow(cleaned_row)
            skip_line = False


remove_and_exclude('data/tenge2.csv', 'data/tenge3.csv')
remove_and_exclude('data/usd2.csv', 'data/usd3.csv')


def delete_columns(input_file, output_file, columns_to_keep):
    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        # Read the header
        header = next(reader)

        # Find the indices of columns to keep
        indices_to_keep = [header.index(column) for column in columns_to_keep if column in header]

        # Write the header with selected columns
        writer.writerow([header[i] for i in indices_to_keep])

        # Write the rows with selected columns
        for row in reader:
            # Check if the row has enough elements
            if len(row) >= max(indices_to_keep) + 1:
                writer.writerow([row[i] for i in indices_to_keep])


delete_columns('data/usd3.csv', 'data/usd_result.csv', ['regions', 'Year 2022'])
delete_columns('data/tenge3.csv', 'data/tenge_result.csv', ['regions', 'Year 2022'])


def merge_csv(tenge_file, usd_file, output_file):
    with open(tenge_file, 'r', encoding='utf-8') as tenge_csv, \
         open(usd_file, 'r', encoding='utf-8') as usd_csv, \
         open(output_file, 'w', encoding='utf-8', newline='') as result_csv:

        tenge_reader = csv.reader(tenge_csv)
        usd_reader = csv.reader(usd_csv)
        result_writer = csv.writer(result_csv)

        # Read headers
        tenge_header = next(tenge_reader)
        usd_header = next(usd_reader)

        # Check if headers match
        if tenge_header != usd_header:
            raise ValueError("Headers of the input CSV files do not match.")

        # Write the header to the result CSV
        result_writer.writerow(['regions', 'Year 2022 ₸', 'Year 2022 $'])

        # Merge the data from both CSVs
        for tenge_row, usd_row in zip(tenge_reader, usd_reader):
            # Check if regions match
            if tenge_row[0] != usd_row[0]:
                raise ValueError("Regions in the input CSV files do not match.")

            # Write the merged row to the result CSV
            result_writer.writerow([tenge_row[0], tenge_row[1], usd_row[1]])


merge_csv('data/tenge_result.csv', 'data/usd_result.csv', 'data/result.csv')


def delete_files(file_paths):
    for file_path in file_paths:
        try:
            os.remove(file_path)
        except FileNotFoundError:
            print(f"File '{file_path}' not found.")
        except Exception as e:
            print(f"Error deleting file '{file_path}': {e}")


files_to_delete = ['data/tenge.csv', 'data/tenge2.csv', 'data/tenge3.csv', 'data/usd0.csv', 'data/usd.csv', 'data/usd2.csv', 'data/usd3.csv','data/tenge_result.csv', 'data/usd_result.csv']
delete_files(files_to_delete)


package = Package()
package.infer(r"data\result.csv")
package.commit()
package.save(r"datapackage.json")
