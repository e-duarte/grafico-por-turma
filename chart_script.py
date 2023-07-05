import matplotlib.pyplot as plt
import pandas as pd
import json
import numpy as np


with open('header_portugues345.json', encoding="utf8") as config_file:
    header = json.load(config_file)

SUBJECT = header['subject']
FICHA_PATH = 'relatório 34.xlsx'
SHEET_NAME = 'PORTUGUÊS 3 ANO'

ficha = pd.read_excel(FICHA_PATH, SHEET_NAME)

def extract_data_classroom(ficha):
    data = {}
    for index, d in ficha.iterrows():
        d.dropna(inplace=True)
        if index >= 11 and d[0] != 'TOTAL':
            d = list(d)
            classroom = d[0]
            data[classroom] = {}
            results = d[1:]

            for vs, vls in zip(header['vars'], header['values']):
                data[classroom][vs] = {'label':[], 'vals':[]}
                for vl in vls:
                    item = results[0]
                    del(results[0])
                    data[classroom][vs]['label'].append(vl)
                    data[classroom][vs]['vals'].append(item)
           
    return data


def remove_zero(data):
    for c in data:
        for v in data[c]:
            if 0 in data[c][v]['vals']:
                i = data[c][v]['vals'].index(0)
                data[c][v]['vals'].remove(0)
                del(data[c][v]['label'][i])

    return data
        
data = extract_data_classroom(ficha)
data = remove_zero(data)

for classroom in data:
    for variable in data[classroom].keys():
        labels = data[classroom][variable]['label']
        vals = data[classroom][variable]['vals']
        print(classroom)

        fig1, ax1 = plt.subplots(figsize=(9, 9))
        p = ax1.pie(vals, autopct='%1.1f%%', startangle=90,  textprops={'fontsize': 19})
        ax1.axis('equal')
        ax1.set_title(variable)
        plt.legend(p[0], labels, loc='best')
        # plt.show()
        fig1.savefig(f'out/{classroom}_{SUBJECT}_{variable}.png', dpi=150)