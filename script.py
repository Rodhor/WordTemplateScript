import pandas as pd
from docxtpl import DocxTemplate
import glob
from pathlib import Path
import os
from datetime import datetime

# Load Table with info
df = pd.read_excel("Test_Infos.xlsx")

# Create output folder if it does not exist
output_folder = Path(f"Outputfolder {datetime.today().strftime('%B %d')}")
output_folder.mkdir(exist_ok=True)

# Create list of all Worddocs
docx_files = glob.glob("*.docx")

# Loop through each row in the infotable
for index, row in df.iterrows():

    # Loop through each file list of worddocs and render the template with the data from the current row
    for file in docx_files:
        doc = DocxTemplate(file)
        doc.render(row.to_dict())

        # Save the file with the Name of the particpant
        doc.save(f"{output_folder}{os.sep}{row['Name']}-{os.path.basename(file)}")
