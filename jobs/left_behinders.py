import matplotlib.pyplot as plt
from operator import itemgetter

from data_loaders import Archive, Client_List, UNT
from figure_generators import PieChartGenerator
from outfile_generators import XLSX_Out

import sys
sys.path[0] = sys.path[0].removesuffix('\jobs')

# Loading and processing data from input Excel files
client_list = Client_List.Client_List(filepath="Client List.xlsx", last_stages=["Government", "Government", "FSP", "FSP"])
archive = Archive.Archive(filepath="ALL Services Archived.xlsx", indices=[0, 1, 2, 3])
unt_list = {"October" : UNT.UNT(filepath="Unique Name Tracker (October).xlsx", indices=[1, 2, 5, 4]),
            "November" : UNT.UNT(filepath="Unique Name Tracker (November).xlsx", indices=[1, 2, 5, 4])}

# Processing data according to rules given by client
for month in unt_list:
    left_behinders = []
    stages = [
        {"In Progress" : 0, "Finished" : 0, "Left Behind" : 0, "Found in Client Archived List" : 0},
        {"In Progress" : 0, "Finished" : 0, "Left Behind" : 0, "Found in Client Archived List" : 0},
        {"In Progress" : 0, "Finished" : 0, "Left Behind" : 0, "Found in Client Archived List" : 0},
        {"In Progress" : 0, "Finished" : 0, "Left Behind" : 0, "Found in Client Archived List" : 0}
    ]
    for i in range(len(client_list.client_list)):
        for j in range(1, len(client_list.client_list[i])):
            row = [*client_list.client_list[i][j],]
            if row[-2] == client_list.last_stages[i] and row[-1] == "Ready":
                stages[i]["Finished"] += 1
            elif row[1].strip() in archive.archive[i]:
                stages[i]["Found in Client Archived List"] += 1
            elif row[1].strip() not in unt_list[month].unt[i]:
                stages[i]["Left Behind"] += 1
                left_behinders.append([row[1].strip(), client_list.names[i]])
            else:
                stages[i]["In Progress"] += 1

    charts = PieChartGenerator.PieChartGenerator(data=stages, titles=client_list.names)   # Pie charts generated and displayed

    left_behinders = sorted(left_behinders, key=itemgetter(0))
    left_behinders.insert(0, ["Client Name", "Service"])

    # Output file in requested format
    outfile = XLSX_Out.XLSX_Out(left_behinders, "Left_Behinders_*.xlsx".replace("*", month))