import json
import pandas as pd

with open('sellers.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for seller in data:
    seller['seller'] = bytes(seller['seller'], 'utf-8').decode('unicode_escape').encode('latin1').decode('utf-8')

df = pd.DataFrame(data)

df.to_excel('sellers.xlsx', index=False)
