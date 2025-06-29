git clone https://github.com/Tomspartan/cloverimport.git

python pip install -r requirements.txt

copy existing main.py to account for source excel document structure. 

within source file python
# Read data from the source Excel file
source_file = "cloverimport/SkechersSummer2025Draft.xlsx"
columns = ['Style #','Style Name','Color','Size','UPC','Cost']
data = pd.read_excel(source_file,header=0,usecols=columns,dtype=object) # 0-indexed
