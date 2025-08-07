import pandas as pd
import openpyxl
import shutil

# Read data from the source Excel file
source_file = 'Spring2026LtApparel(Kids Carhartt).xlsx'
columns = ['Style','Color','Size','UPC Num','cost','Our Price ']
data = pd.read_excel(source_file,header=0,usecols=columns) # 0-indexed

# Drop rows that are Carryover


#Replace NaN with empty strings in the following columns. Needed for name creation. 
data['Style'] = data['Style'].fillna('')
data['Color'] = data['Color'].fillna('')
data['Size'] = data['Size'].fillna('')

# Create Marc approved name for use on customer receipt
data['Name'] = data[['Style','Color','Size']].aggregate('-'.join, axis=1)
data['Name'] = data['Name'].str.rstrip('-')

# Remove the now unneeded columns
data.drop(columns={'Style','Color','Size'},inplace=True)

# Rename UPC column to Product Code in order to match Clover format
data.rename(columns={'UPC Num':'Product Code'},inplace=True)
data.rename(columns={'Our Price':'Price Unit'},inplace=True)
# Calculate Price unit based on value in Cost column.


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
data.insert(11,'Hidden','')
data.insert(12,'Non-revenue item','')

# Copy the template file to the output file. We'll overwrite the included 'Items' sheet in the next step.
template = 'CloverInventoryTemplate_Small.xlsx'
output_file = 'CloverInventoryoutputSpringKids26.xlsx'
shutil.copy2(template,output_file)

# Write the data to the output file, in apend mode, replacing the only the 'Items' sheet
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    data.to_excel(writer,sheet_name='Items', index=False)