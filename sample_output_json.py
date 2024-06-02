import pandas as pd

# Replace 'path_to_your_file.csv' with the actual file path
csv_file_path = 'sample-output.csv'

# Load the CSV file
data = pd.read_csv(csv_file_path)

# Convert to JSON
data_json = data.to_json(orient='records')

# Print the JSON data
print(data_json)

# Optionally, save the JSON data to a file
with open('output.json', 'w') as json_file:
    json_file.write(data_json)