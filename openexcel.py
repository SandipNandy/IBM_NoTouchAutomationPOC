import os
from pathlib import Path
# opening EXCEL through Code
                    #local path in dir
absolutePath = Path('./Statusstamping for CA.xlsx').resolve()
os.system(f'start excel.exe "{absolutePath}"')