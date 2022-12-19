import xlwings as xw
import openpyxl as xl
import pandas as pd
class information:

    def __init__(self, database):
        self.database = xw.Book(database)

    def get_dataset_column(self, sheet, data_range):
        return self.database.sheets[sheet].range(data_range).value


source_database_path = input("Please input path for input: ")
# output_database_path = input("Please input path for output: ")
source_database = information(source_database_path)
# Database Created
output_database = pd.DataFrame()


# for _ in range(1): #Number of sheets
current_sheet_name = input("Please input sheet name: ")
how_many_data = input("How many columns will you select from this sheet?: ")
for i in range(int(how_many_data)): # Get info of specific column
    data_name = input("Please input what data you are selecting: ")
    range_data = input("Please input range for " + data_name + ": ")
    data = source_database.get_dataset_column(current_sheet_name, range_data)
    output_database.insert(i, data_name, data)
    print("Done inserting " + data_name + " column, ready for next.......................")


output_database.to_csv('pandas/' + current_sheet_name + ".csv", index= False)
print(output_database)

