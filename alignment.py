from openpyxl import Workbook, load_workbook
import os
import datetime


def setup():
    column_counter = 2
    last_ws_row = 2

    return column_counter, last_ws_row


def mock_data(ws):
    pass
    # MOCK DATA for ws
    # ws['A2'] = 'CAASNTDKLIF'
    # ws['A3'] = 'CAAKAGGTSYGKLTF'
    # ws['A4'] = 'CAASDSWGKLQF'
    # ws['A5'] = 'CAANTNAGKSTF'
    # ws['A6'] = 'CAASTGRRALTF'
    # ws['D1'] = 'MNFAKER'
    # ws['B2'] = 1
    # ws['B3'] = 5
    # ws['B4'] = 10
    # ws['B5'] = 3
    # ws['B6'] = 4


# def old_slow_wb_update(current_ws, current_wb, last_ws_row):
#     # Iterate through rows of current worksheet
#     for row in current_ws.iter_rows(min_row=3, min_col=1, max_col=6):
#         aa_seq = None
#         aa_seq_count = None
#         found_in_wb = False
#
#         for cell in row:
#             if cell.column == 'A':
#                 # Assign amino acid sequence variable for current row
#                 aa_seq_count = cell.value
#             elif cell.column == 'F':
#                 # Assign sequence count variable for current row
#                 aa_seq = cell.value
#
#         print(f'{aa_seq}: {aa_seq_count}')
#
#         # Iterate through rows of joined worksheet and update if found
#         for ws_row in ws.iter_rows(
#                 min_row=2, min_col=1, max_col=column_counter):
#
#             ws_aa_seq = ws_row[0].value
#             if ws_aa_seq == aa_seq:
#                 found_in_wb = True
#                 print(f"match for {aa_seq} at {current_wb.sheetnames[0]}, "
#                       f"{cell.column}{cell.row}")
#
#                 ws_row_value = ws_row[column_counter - 1].value
#                 print(f'before: {ws_row_value}')
#
#                 print(f'adding aa_seq: {aa_seq_count}')
#
#                 if ws_row_value:
#                     ws_row_value += aa_seq_count
#                 else:
#                     ws_row_value = aa_seq_count
#
#                 ws.cell(row=ws_row[0].row, column=column_counter,
#                         value=ws_row_value)
#
#                 print(f'after: {ws_row_value}')
#
#             else:
#                 pass
#                 # print("HERE")
#                 # if not ws[column_counter-1]:
#                 #     ws.cell(row=ws_row[0].row, column=column_counter,
#                 #             value=0)
#
#         if not found_in_wb:
#             print("adding new row to wb...")
#             ws.cell(row=last_ws_row, column=1, value=aa_seq)
#             ws.cell(row=last_ws_row, column=column_counter,
#                     value=aa_seq_count)
#             last_ws_row += 1

def create_dict():
    wb_dict = dict()
    return wb_dict

def get_file_count():
    file_count = 0
    directory_path = os.path.dirname(os.path.realpath(__file__))
    for file in os.listdir(directory_path):
        filename = os.fsdecode(file)
        if filename.endswith(".xlsx") and not filename.startswith(
                ('~', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')):
            file_count += 1

    return file_count


def iterate_over_files(column_counter, last_ws_row):

    file_count = get_file_count()
    wb_dict = create_dict()

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
            print(current_wb.sheetnames[0])
            print('***********************************************************')

            # # First sheet name in file will be the name of the header
            # ws.cell(row=1, column=column_counter,
            #         value=current_wb.sheetnames[0])

            # Iterate through rows of current worksheet
            for row in current_ws.iter_rows(min_row=3, min_col=1, max_col=6):
                aa_seq = None
                aa_seq_count = None
                found_in_wb = False

                for cell in row:
                    if cell.column == 'A':
                        # Assign amino acid sequence variable for current row
                        aa_seq_count = cell.value
                    elif cell.column == 'F':
                        # Assign sequence count variable for current row
                        aa_seq = cell.value

                print(f'{aa_seq}: {aa_seq_count}')

                



            column_counter += 1
            continue
        else:
            continue


def read_file():
    pass


def save_workbook():
    # https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
    wb = Workbook()
    ws = wb.active
    ws.title = "combined_matrix"

    ws['A1'] = 'Motifs'

    today_date = datetime.datetime.date(datetime.datetime.now())
    new_file_name = str(today_date) + '_joined_excel_data.xlsx'
    wb.save(new_file_name)
    print("\n ******* FILE SAVED *********\n")


if __name__ == '__main__':
    column_counter, last_ws_row = setup()

    iterate_over_files(column_counter, last_ws_row)

    # save_workbook()
