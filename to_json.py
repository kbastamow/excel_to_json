import pandas as pd
import json

def convert_excel_csv_to_json(input_excel, output_json, sheet_name, start_row, finish_row=None, key_column=0, value_column=1):
    df = pd.read_excel(input_excel, sheet_name=sheet_name, header=None)  # Use read_csv if your file is in CSV format
   
    sliced_df = df.iloc[start_row-1:finish_row]

    # Delete accidental line breaks in the DataFrame
    sliced_df = sliced_df.replace({1: {r'\r': ' ', r'\n': ''}}, regex=True)

    # Replace NaN values with an empty string - they correspond to empty cells
    sliced_df = sliced_df.fillna("")

    nested_dict = {}
    current_parent = None

    for index, row in sliced_df.iterrows():
        key = row[key_column]
        value = row[value_column]

        # If the value is empty, it becomes the parent for the next values until the next gap
        if not value:
            current_parent = key
            nested_dict[current_parent] = {}
        elif current_parent:
            nested_dict[current_parent][key] = value

    with open(output_json, 'w', encoding='utf-8') as json_file:
        json.dump(nested_dict, json_file, ensure_ascii=False, indent=2)

# Specify your arguments
convert_excel_csv_to_json("Mock_excel.xlsx", "output.json", "Home")


