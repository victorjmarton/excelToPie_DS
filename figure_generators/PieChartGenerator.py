import matplotlib.pyplot as plt
import os

def autopct_format(data):
    def display_format(pct):
        total = sum(data)
        val = int(round(pct*total/100.0))
        return '{v:d}'.format(v=val)
    return display_format

class PieChartGenerator:
    def __init__(self, data, titles):
        self.data = data
        self.titles = titles
        self.generate_charts()

    def generate_charts(self):
        for i in range(len(self.data)):
            _, ax = plt.subplots()
            ax.pie(self.data[i].values(), autopct=autopct_format(self.data[i].values()))
            ax.set_title(self.titles[i])

            plt.legend(labels=self.data[i].keys(), loc="best", fontsize=10)
        
        plt.tight_layout()
        plt.show()