import pandas as pd
import json

def convert_json_to_excel(input_json, output_excel, sheet_name="Sheet 1", start_row=1, key_column=0, value_column=1):
    with open(input_json, 'r', encoding='utf-8') as json_file:
        data = json.load(json_file)

    # Create an empty DataFrame
    df = pd.DataFrame(columns=[key_column, value_column])

    row_index = start_row-1

    # Populate the DataFrame
    for parent_key, child_dict in data.items():
        df.loc[row_index, key_column] = parent_key
        df.loc[row_index, value_column] = ''
        row_index += 1

        for child_key, child_value in child_dict.items():
            df.loc[row_index, key_column] = child_key
            df.loc[row_index, value_column] = child_value
            row_index += 1

    # Write the DataFrame to an Excel file
    df.to_excel(output_excel, sheet_name=sheet_name, header=None, index=False)

# Specify your arguments
convert_json_to_excel('Mock_json.json', 'output.xlsx', "Test")