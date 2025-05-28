# Clean the Superstore data
# Add a column with the number of days it takes between order and ship dates
# Convert State names to abbreviations
# Remove the Country column since all data is from the US only
import pandas as pd
from datetime import datetime

# Read the Superstore data into a dataframe
superstore_data_location = ("D:\_GitHubRepo\Python\SuperStoreDataCleaning\DataFiles\SuperstoreEncoded_05142025.csv")
superstore_data_location_output = ("D:\_GitHubRepo\Python\SuperStoreDataCleaning\DataFiles\Superstore_DataCleaning.xlsx")
superstore_data_location_output_duplicate = ("D:\_GitHubRepo\Python\SuperStoreDataCleaning\DataFiles\Superstore_DataCleaning_Duplicates.xlsx")
dataframe_superstore = pd.read_csv(superstore_data_location)
#pd.reset_option('display.max_columns', None)
# print(dataframe_superstore)

# Take a look at the dataframe info to see if there are nulls and to see the data types for the columns
print(dataframe_superstore.info())

# Get the number of days between order and ship dates
dataframe_superstore["Order Date"] = dataframe_superstore["Order Date"].astype("string")
dataframe_superstore["Ship Date"] = dataframe_superstore["Ship Date"].astype("string")
order_date = pd.to_datetime(dataframe_superstore["Order Date"])
ship_date = pd.to_datetime(dataframe_superstore["Ship Date"])
days_to_ship = (ship_date - order_date).dt.days.astype(int)

# Add a column that will display how many days between order date and ship date
dataframe_superstore.insert(4, "Days To Ship",  days_to_ship)
pd.set_option('display.max_columns', None)

# Create a dictionary with state names and their abbreviations
state_abbreviations = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "District of Columbia": "District of Columbia", "Florida": "FL", "Georgia": "GA", 
    "Hawaii": "HI", "Idaho": "ID", "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", 
    "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK",
    "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY"
}

# Convert state names to abbreviations
dataframe_superstore["State"] = dataframe_superstore["State"].map(state_abbreviations)

# Remove the Country column
del dataframe_superstore["Country"]

# There are products that share the same product id therefor not making the product Id's unique.
# The product id's will be corrected by finding the products that share the same product ids and updating the product id of the product containing the duplicate. The product id will increase by 1 if the id doesn't already exist.  If it does exist when incrementing then we will check if there is an existing id when we decrement by 1 and update the id if it doesn't.

# Use the regex library
import re 

# Sort by ProductID and ProductName
dataframe_superstore_sorted = dataframe_superstore.sort_values(by=['Product ID', 'Product Name'])

# Find Product IDs with more than one unique Product Name
duplicates = dataframe_superstore_sorted.groupby('Product ID').filter(lambda x: x['Product Name'].nunique() > 1)


duplicates.to_excel(superstore_data_location_output_duplicate, index=False)

# Function to adjust duplicate Product IDs and propagate changes
def adjust_product_ids(df):
    # Track existing IDs
    existing_ids = set(df['Product ID'])  
    
    # Store mapping for modified Product Names
    product_name_to_new_id = {}  
    product_groups = df.groupby('Product ID')

    for product_id, group in product_groups:
        unique_names = group['Product Name'].unique()
        
        if len(unique_names) > 1:
            # Preserve the first occurrence as the original
            original_name = unique_names[0]

            for index, row in group.iterrows():
                if row['Product Name'] != original_name:
                    match = re.search(r'(\d+)', row['Product ID'])
                    
                    if not match:
                        continue  # Skip if no numeric portion
                    
                    numeric_part = int(match.group(1))
                    
                    # Extract prefix
                    prefix = row['Product ID'].replace(str(numeric_part), "")  
                    
                    # Try incrementing first
                    new_id = f"{prefix}{numeric_part + 1}"

                    if new_id not in existing_ids:
                        print(f'New ID + 1: {new_id}')
                        print(f'Row Product Name: {row['Product Name']} Original Name: {original_name}')
                        
                        
                        df.at[index, 'Product ID'] = new_id
                        existing_ids.add(new_id)
                        product_name_to_new_id[row['Product Name']] = new_id
                    else:
                        # If incremented ID exists, try decrementing
                        new_id = f"{prefix}{numeric_part - 1}"
                        
                        if new_id not in existing_ids and numeric_part - 1 > 0:
                            print(f'New ID - 1: {new_id}')
                            print(f'Row Product Name: {row['Product Name']} Original Name: {original_name}')
                        
                            df.at[index, 'Product ID'] = new_id
                            existing_ids.add(new_id)
                            product_name_to_new_id[row['Product Name']] = new_id

    # Propagate updated Product IDs for all matching names
    df['Product ID'] = df['Product Name'].map(lambda name: product_name_to_new_id.get(name, df.loc[df['Product Name'] == name, 'Product ID'].iloc[0]))

    return df

# Apply the adjustment function
dataframe_superstore_updated = adjust_product_ids(dataframe_superstore_sorted)

# Write dataframe back to .xlsx
dataframe_superstore_updated.to_excel(superstore_data_location_output, index=False)