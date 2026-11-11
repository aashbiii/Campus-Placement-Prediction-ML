import json
import os

with open('Campus_Placement_Prediction.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

print(f"Total cells: {len(nb['cells'])}")
for i, cell in enumerate(nb['cells']):
    ctype = cell['cell_type']
    source = "".join(cell.get('source', []))
    first_line = source.split('\n')[0] if source else ''
    outputs = cell.get('outputs', [])
    out_types = [o.get('output_type', '') for o in outputs]
    has_image = any('image/png' in o.get('data', {}) for o in outputs if 'data' in o)
    print(f"Cell {i:02d} ({ctype}): {first_line[:60]} | outputs: {out_types} | img: {has_image}")
