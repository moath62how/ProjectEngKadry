from pathlib import Path
import pandas as pd
from src.excel_handler import read_national_ids_from_excel

p = Path('test_arabic_ids.xlsx')
df = pd.DataFrame({'ANAME':['أحمد صبري','سارة علي'],'الرقم القومى':['29709101300615','29501011234567']})
df.to_excel(p,index=False)
print('Wrote', p)
ids = read_national_ids_from_excel(str(p))
print('Detected IDs:', ids)
