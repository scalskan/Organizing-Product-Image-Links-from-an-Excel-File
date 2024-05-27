import pandas as pd

# Read the Excel file
df = pd.read_excel("Photos.xlsx")

# Count the number of times each code is repeated
code_repeat_count = df['Option'].value_counts().to_dict()

# Create an empty list to store the new information
new_data = []

# Verify the column names. There are 7 columns with information, so we list 7 columns.
column_names = ['Web Image Link1', 'Web Image Link2', 'Web Image Link3', 'Web Image Link4', 'Web Image Link5', 'Web Image Link6', 'Web Image Link7']

# Iterate through each row
for index, row in df.iterrows():
    style_code = row['Style Code']  # Get the Style Code column
    code = row['Option']  # Get the Options column
    repeat_count = code_repeat_count.get(code, 0)  # Get the repeat count of the code

    # Check each column and add if there is an image link
    for column in column_names:
        if column in row and pd.notna(row[column]):
            new_entry = {
                'Style Code': style_code,
                'Code': code.strip(),
                'Image Link': row[column].strip(),
                'Repeat Count': repeat_count
            }
            new_data.append(new_entry)

# Create a DataFrame containing the new information
new_df = pd.DataFrame(new_data)

# Write to a new Excel file
new_df.to_excel("New_Data.xlsx", index=False)
