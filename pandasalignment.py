# # https://openpyxl.readthedocs.io/en/stable/
# from openpyxl import load_workbook
# https://xlsxwriter.readthedocs.io/
# import xlsxwriter
import os
import datetime
import pandas as pd


def get_file_count():
    """Gets the number of files to read and creates list from them"""

    header_list = []
    file_list = []
    file_count = 0

    # https://stackoverflow.com/questions/10377998/how-can-i-iterate-over-files-in-a-given-directory
    # This gives the directory path from which the .py file is being run
    directory_path = os.path.dirname(os.path.realpath(__file__))
    for file in os.listdir(directory_path):
        filename = os.fsdecode(file)
        # Append not temporary (~) excel files to header_list
        if filename.endswith(".xlsx") and not filename.startswith('~'):
            file_count += 1
            header_list.append(filename.strip('.xlsx'))
            file_list.append(filename)

    print()
    print(f"file count is {file_count}")
    print(f"files working with:")
    print(file_list)
    print()

    return file_count, file_list, header_list, directory_path


def iterate_over_files(file_count, file_list, directory_path):
    """Reads excel files, adds data-frames of unique aa seqs to list"""

    df_list = []


    # Iterate over each file in the current folder
    for filename in file_list:
        header_name = filename.strip('.xlsx')

        # Open current workbook and go to its worksheet for reading
        current_file_path = os.path.join(directory_path, filename)
        print(f"\n******* Opening {filename} *******")
        df = pd.read_excel(current_file_path)
        df.columns = df.iloc[0]
        df = df.reindex(df.index.drop(0))
        df = df[['Read count', 'CDR3 amino acid sequence']]
        # print(df)
        sum_df = df.groupby('CDR3 amino acid sequence').sum().sort_values(
            'Read count', ascending=False)
        sum_df.columns = [header_name]
        # sum_df.reset_index('CDR3 amino acid sequence', inplace=True)
        # print(sum_df)

        df_list.append(sum_df)


    print(df_list)
    # return df_list


def combine_data_frames(df_list):
    pass



def write_to_workbook(wb_dict, header_list, is_test=False):
    """Writes a .xlsx excel workbook from the dictionary given and saves"""

    # Get today's date for naming purposes
    today_date = datetime.datetime.date(datetime.datetime.now())
    if is_test:
        new_file_name = "1test_" + str(today_date) + "_combined_matrix.xlsx"
    else:
        new_file_name = str(today_date) + "_combined_matrix.xlsx"

    # Name excel workbook and excel worksheet
    workbook = xlsxwriter.Workbook(new_file_name)
    worksheet = workbook.add_worksheet("combined_matrix")

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    print("\n******* WRITING EXCEL WORKBOOK *********\n")
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

    # Close (and thus save) excel workbook
    workbook.close()

    print("\n************** FILE SAVED **************\n")


if __name__ == '__main__':
    pass

    file_count, file_list, header_list, directory_path = get_file_count()

    df_list = iterate_over_files(file_count, file_list, directory_path)

    blah = combine_data_frames(df_list)

    # print(wb_dict)
    #
    # write_to_workbook(wb_dict, header_list)
