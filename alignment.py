from openpyxl import load_workbook
import os
import datetime


def get_file_count():
    header_list = []
    file_count = 0
    directory_path = os.path.dirname(os.path.realpath(__file__))
    for file in os.listdir(directory_path):
        filename = os.fsdecode(file)

        if filename.endswith(".xlsx") and not filename.startswith(
                ('~', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
            file_count += 1
            header_list.append(filename.strip('.xlsx'))

    print(f"file count is {file_count}")
    print(f"files working with:")
    print(header_list)

    return file_count, header_list


def iterate_over_files(file_count):
    column_counter = 0
    wb_dict = dict()

    # https://stackoverflow.com/questions/10377998/how-can-i-iterate-over-files-in-a-given-directory
    # This gives the directory path from which the .py file is being run
    directory_path = os.path.dirname(os.path.realpath(__file__))
    for file in os.listdir(directory_path):
        filename = os.fsdecode(file)
        if filename.endswith(".xlsx") and not filename.startswith(
                ('~', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
            current_file_path = os.path.join(directory_path, filename)
            current_wb = load_workbook(current_file_path)
            current_ws = current_wb.active

            print('***********************************************************')
            print(f"working on {current_wb.sheetnames[0]}")
            print('***********************************************************')

            # # First sheet name in file will be the name of the header
            # ws.cell(row=1, column=column_counter,
            #         value=current_wb.sheetnames[0])

            # Iterate through rows of current worksheet
            for row in current_ws.iter_rows(min_row=3, min_col=1, max_col=6):
                aa_seq = None
                aa_seq_count = None

                for cell in row:
                    if cell.column == 'A':
                        # Assign amino acid sequence variable for current row
                        aa_seq_count = cell.value
                    elif cell.column == 'F':
                        # Assign sequence count variable for current row
                        aa_seq = cell.value

                # print(f'{aa_seq}: {aa_seq_count}')

                # Add amino acid sequence to workbook and populate all columns 0
                if aa_seq not in wb_dict:
                    wb_dict[aa_seq] = [0 for x in range(file_count)]

                # Add amino acid sequence count to current column of dict
                wb_dict[aa_seq][column_counter] += aa_seq_count

            # Move to next column
            column_counter += 1
            pass
        else:
            pass

    return wb_dict


def write_to_workbook(wb_dict, header_list):
    # https://xlsxwriter.readthedocs.io/
    import xlsxwriter

    today_date = datetime.datetime.date(datetime.datetime.now())
    new_file_name = str(today_date) + "_joined_excel_data.xlsx"

    workbook = xlsxwriter.Workbook(new_file_name)
    worksheet = workbook.add_worksheet("combined_matrix")

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    print("\n ******* WRITING EXCEL WORKBOOK *********\n")
    # Add all headers to excel file
    worksheet.write(row, col, 'Motifs')
    col += 1
    for header in header_list:
        worksheet.write(row, col, header)
        col += 1


    # write dictionary to excel file
    for aa_seq in wb_dict:
        col = 0
        row += 1
        worksheet.write(row, col, aa_seq)
        for aa_seq_count in wb_dict[aa_seq]:
            col += 1
            worksheet.write(row, col, aa_seq_count)

    workbook.close()
    print("\n ************ FILE SAVED ***************\n")


if __name__ == '__main__':

    file_count, header_list = get_file_count()

    wb_dict = iterate_over_files(file_count)

    write_to_workbook(wb_dict, header_list)
