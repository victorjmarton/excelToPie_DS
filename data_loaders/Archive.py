import openpyxl

class Archive:
    def __init__(self, filepath, indices):
        self.filepath = filepath
        self.indices = indices
        self.archive = self.load_archive()

    # Loads and formats the Archive file
    def load_archive(self):
        archive = openpyxl.load_workbook(self.filepath)
        return self.reformat(archive)

    # Converts data from the .xlsx file into a workable list in the specified sheet order
    def reformat(self, worksheet : openpyxl.workbook.workbook.Workbook):
        sheets = [[], [], [], []]

        current = 0
        for i in self.indices:
            for tuple in list(worksheet[worksheet.sheetnames[i]].values):
                row = [*tuple,]
                if row[1] is not None:
                    sheets[current].append(row[1].strip())
            current += 1
        return sheets