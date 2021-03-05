import pandas as pd
import json
import os

excel_file = 'System Colours.xlsx'
excel_sheets = ["Pipes", "Ducts", "Trays", "Plant Zones", "Service Zones"]

for sheet in excel_sheets:

    output_json_file = sheet + '.json'
    output_list = []

    df = pd.read_excel( os.path.join( os.getcwd(), excel_file ), sheet_name = sheet).dropna( axis='columns' )

    for index, row in df.iterrows():

        json_row = {
            "system_name" : row[0],
            "fill" : {
                "R": row['R'],
                "G": row['G'],
                "B": row['B'],
            },
            "outline": {
                "R": row['R.1'] if 'R.1' in row else row['R'],
                "G": row['G.1'] if 'G.1' in row else row['G'],
                "B": row['B.1'] if 'B.1' in row else row['B'],
            },
        }
        output_list.append(json_row)

    with open(os.path.join(os.getcwd(), output_json_file), 'w') as output_handle:
        json.dump(output_list , output_handle, indent = 4)