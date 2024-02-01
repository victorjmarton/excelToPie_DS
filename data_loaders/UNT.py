import openpyxl

class UNT:
    def __init__(self, filepath, indices):
        self.filepath = filepath
        self.indices = indices
        self.unt = self.load_UNT()

    # Loads and formats the Unique Name Tracker file
    def load_UNT(self):
        unt = openpyxl.load_workbook(self.filepath)
        return self.reformat(unt)

    # Converts data from the .xlsx file into a workable list in the specified sheet order
    def reformat(self, worksheet : openpyxl.workbook.workbook.Workbook) -> list:
        sheets = [[], [], [], []]
        
        current = 0
        for i in self.indices:        
            for tuple in list(worksheet[worksheet.sheetnames[i]].values):
                row = [*tuple,]
                if row[0] is not None:
                    sheets[current].append(row[0].strip())
                if row[3] is not None:
                    sheets[current].append(row[3].strip())
            current += 1
        return sheets