import pandas as pd
from functools import reduce
import os
import datetime


def get_file_count():
    """Gets the number of files to read and creates list from them"""

    file_list = []
    file_count = 0
    # https://stackoverflow.com/questions/10377998/how-can-i-iterate-over-files-in-a-given-directory
    # This gives the directory path from which the .py file is being run
    directory_path = os.path.dirname(os.path.realpath(__file__))
    for file in os.listdir(directory_path):
        filename = os.fsdecode(file)

        if filename.endswith(".xlsx") and not filename.startswith('~'):
            file_count += 1
            file_list.append(filename)

    print()
    print(f"file count is {file_count}")
    print(f"files working with:")
    print(file_list)
    print()

    return file_count, file_list, directory_path


def open_files(file_list, directory_path):
    """Opens workbooks, gets count of and list of headers"""

    df_list = []

    # Iterate over each file in the current folder
    for filename in file_list:
        # Open current workbook and go to its worksheet for reading
        print(f"\n******* Opening {filename} *******")
        current_file_path = os.path.join(directory_path, filename)
        current_df = pd.read_excel(current_file_path)
        current_df.set_index(current_df.columns[0], inplace=True)
        df_list.append(current_df)

    return df_list


def combine_data_frames(dfs):
    """joins dataframes and cleans them up"""

    print("\n******* Combining workbooks *********\n")

    # http://notconfusing.com/joining-many-dataframes-at-once-in-pandas-n-ary-join/
    def join_dfs(ldf, rdf):
        return ldf.join(rdf, how='outer')

    # Combine data frames
    final_df = reduce(join_dfs, dfs)
    # Fill NaN values with 0
    final_df.fillna(value=0, inplace=True)
    # Sort columns high to low by workbook
    final_df.sort_values(
        [final_df.columns[x] for x in range(len(final_df.columns))],
        ascending=False, inplace=True)
    # print([final_df.columns])

    # Convert float values to ints
    final_df = final_df.astype(int)

    return final_df


def write_to_workbook(combo_df):
    """Writes a .xlsx excel workbook from the dictionary given and saves"""

    # Get today's date for naming purposes
    today_date = datetime.datetime.date(datetime.datetime.now())
    file_name = str(today_date) + "_multi_matrix.xlsx"
    sheet_name = "multi_matrix"

    print("\n******* WRITING EXCEL WORKBOOK *********\n")

    combo_df.to_excel(file_name, sheet_name=sheet_name)

    print("\n************** FILE SAVED **************\n")


if __name__ == '__main__':
    file_count, file_list, directory_path = get_file_count()
    df_list = open_files(
        file_list, directory_path)

    combo_df = combine_data_frames(df_list)
    # print(combo_df)
    write_to_workbook(combo_df)
