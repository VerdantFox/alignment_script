from openpyxl import Workbook, load_workbook
import openpyxl
import os
import datetime




def read_file():
    pass

if __name__ == '__main__':

    # https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
    wb = Workbook()
    ws = wb.active
    ws.title = "combined_matrix"

    ws['A1'] = 'Motifs'

    column_counter = 2
    row_counter = 2

    today_date = datetime.datetime.date(datetime.datetime.now())

    # MOCK DATA for ws
    ws['A2'] = 'CAASNTDKLIF'
    ws['A3'] = 'CAAKAGGTSYGKLTF'
    ws['A4'] = 'CAASDSWGKLQF'
    ws['A5'] = 'CAANTNAGKSTF'
    ws['A6'] = 'CAASTGRRALTF'
    ws['B1'] = 'MNFAKER'
    ws['B2'] = 1
    ws['B3'] = 5
    ws['B4'] = 10
    ws['B5'] = 3
    ws['B6'] = 4



    # print(wb.sheetnames)
    # c = ws['A4']
    # ws['A4'] = 4
    # print(c)
    # print(ws['A4'])
    #
    # for row in ws.iter_rows(min_row=1, max_col=3, max_row=2):
    #     for cell in row:
    #         print(cell)
    #
    # print(ws['A4'].value)
    # print(c.value)

    # https://stackoverflow.com/questions/10377998/how-can-i-iterate-over-files-in-a-given-directory
    directory_in_string = 'C:/Users/willithe/Desktop/Twin_Alignment/alignment_script'
    directory = os.fsencode('C:/Users/willithe/Desktop/Twin_Alignment/alignment_script')

    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".xlsx") and not \
                filename.startswith(('~', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
            directory_og = os.path.join(directory_in_string, filename)
            new_directory = os.path.normpath(directory_og)
            print(new_directory)
            current_wb = load_workbook(new_directory)
            print(current_wb.sheetnames)
            current_ws = current_wb.active

            # First sheetname is the name of the header
            header_cell = ws.cell(row=1, column=column_counter,
                                  value=current_wb.sheetnames[0])
            for row in current_ws.iter_rows(min_row=3, min_col=1, max_col=6):
                # print(row)
                aa_seq = None
                aa_seq_count = None
                for cell in row:
                    if cell.column == 'A':
                        aa_seq_count = cell.value
                    if cell.column == 'F':
                        aa_seq = cell.value

                for ws_row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                    if ws_row[0].value == aa_seq:
                        print(f"match for {aa_seq} at {current_wb.sheetnames[0]}, {cell.column}{cell.row}")
                        # TODO figure out efficient way to combine all equivalent aa_seq into 1, adding up all the aa_seq_counts

                # print(f'{aa_seq}: {aa_seq_count}')

            column_counter += 1
            continue
        else:
            continue

    new_file_name = str(today_date) + '_joined_excel_data.xlsx'
    wb.save(new_file_name)
