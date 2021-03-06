# By Steven Au
# Motivation: Converts an Excel file with multiple sheets into a JSON file.
import pathlib as pl
import json
import pandas as pd


class ExcelJson:
    """ Converts a standard Excel file into a JSON object. For use with quick conversions."""
    
    def __init__(self, path=None):
        self.path = path
        self.excel_file = {}

    def get_raw_path(self):
        """ Get full path."""
        return self.path

    def get_parent_path(self):
        """ Get parent path name (Excludes the file name + extension)"""
        return pl.Path(self.path).parent

    def get_file_stem(self):
        """ Get file name"""
        return pl.Path(self.path).stem

    def get_parse_excel(self):
        """ PArse the excel."""
        return pd.ExcelFile(self.path)

    def set_path(self):
        """ Set the path to the excel file"""
        self.path = input("Path to Excel: ")

    def excel_magic(self):
        """ Parse the Excel File and save as a dictionary."""
        sheet = 0  # Permits reparsing
        
        excel_object = {}

        try:
            while True:
                sheet_name = self.get_parse_excel().sheet_names[sheet]
                excel_object[sheet_name] = self.get_parse_excel().parse(sheet_name).fillna("NaN").to_dict()
                sheet += 1
        except:
            # File name is the base dictionary => encapsulates the rest of the excel data dictionary
            self.excel_file[self.get_file_stem()] = excel_object
            print("Excel Parsed.")

    def json_export(self):
        """ Export JSON file """
        json_object = json.dumps(self.excel_file)
        json_path = self.get_parent_path().joinpath(self.get_file_stem() + ".json")

        with open(json_path, "w", encoding='utf-8') as json_file:
            json_file.write(json_object)

    def operation_sequence(self):
        """ Linear sequence of operations."""
        if self.path is None:
            # In the event that the user had not initialized with a direct path, then set the path.
            self.set_path()
            
        self.excel_magic()
        self.json_export()


if __name__ == "__main__":
    E2J = ExcelJson()
    E2J.operation_sequence()
