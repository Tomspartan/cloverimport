import pandas as pd
import openpyxl
import shutil

# Read data from the source Excel file
source_file = 'EdREhmann_S25ProductFilterwCost.xlsx'
columns = ['Name','Material','Cost','Size','Dimension','UPC','New/Carryover']
data = pd.read_excel(source_file,header=1,usecols=columns) # 0-indexed

# Drop rows that are Carryover
carryover = data[data['New/Carryover'] == 'Carryover'].index
data.drop(carryover, inplace=True)

#Replace NaN with empty strings in the following columns. Needed for name creation. 
data['Dimension'] = data['Dimension'].fillna('')
data['Size'] = data['Size'].fillna('')
data['Material'] = data['Material'].fillna('')

# Create Marc approved name for use on customer receipt
data['Name'] = data[['Material','Size','Dimension']].aggregate('-'.join, axis=1)
data['Name'] = data['Name'].str.rstrip('-')

# Remove the now unneeded columns
data.drop(columns=['Material','Size','Dimension','New/Carryover'],inplace=True)

# Rename UPC column to Product Code in order to match Clover format
data.rename(columns={'UPC':'Product Code'},inplace=True)

# Calculate Price unit based on value in Cost column.
data['Price Unit'] = (data['Cost']*2)*.85

# Add columns to match Clover format with default or blank values
data.insert(0,'Clover ID','')
data.insert(2,'Alternate Name','')
data.insert(3,'Price','0')
data.insert(4,'Price Type','FIXED')
data.insert(6,'Tax Rates','DEFAULT')
data.insert(7,'SKU','')
data.insert(8,'Modifier Groups','')
data.insert(9,'Quantity','0')
data.insert(10,'Printer Labels','')
data.insert(11,'Hidden','FALSE')
data.insert(12,'Non-revenue item','FALSE')

# Copy the template file to the output file. We'll overwrite the included 'Items' sheet in the next step.
template = 'CloverInventoryTemplate_Small.xlsx'
output_file = 'CloverInventoryoutput.xlsx'
shutil.copy2(template,output_file)

# Write the data to the output file, in apend mode, replacing the only the 'Items' sheet
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer,sheet_name='Items', index=False)