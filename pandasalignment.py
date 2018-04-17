# # https://openpyxl.readthedocs.io/en/stable/
# from openpyxl import load_workbook
# https://xlsxwriter.readthedocs.io/
# import xlsxwriter
import os
import datetime
import pandas as pd
from functools import reduce


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

    # print(df_list)
    return df_list


def combine_data_frames(dfs):
    # http://notconfusing.com/joining-many-dataframes-at-once-in-pandas-n-ary-join/
    def join_dfs(ldf, rdf):
        return ldf.join(rdf, how='outer')

    # Combine data frames
    final_df = reduce(join_dfs, dfs)
    # Fill NaN values with 0
    final_df.fillna(value=0, inplace=True)
    # Sort columns high to low by workbook
    final_df.sort_values([df.columns[0] for df in dfs], ascending=False,
                         inplace=True)
    # Convert float values to ints
    final_df = final_df.astype(int)

    return final_df


def write_to_workbook(final_df):
    """Writes a .xlsx excel workbook from the dictionary given and saves"""

    # Get today's date for naming purposes
    today_date = datetime.datetime.date(datetime.datetime.now())
    file_name = str(today_date) + "_combined_matrix.xlsx"
    sheet_name = "combined_matrix"

    final_df.to_excel(file_name, sheet_name=sheet_name)

    print("\n************** FILE SAVED **************\n")


if __name__ == '__main__':
    pass

    file_count, file_list, header_list, directory_path = get_file_count()

    df_list = iterate_over_files(file_count, file_list, directory_path)

    final_df = combine_data_frames(df_list)

    write_to_workbook(final_df)
