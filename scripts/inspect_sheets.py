import openpyxl
wb = openpyxl.load_workbook(r'output/412YZ/analysis_412YZ.xlsx', data_only=True)

for sheet in ['07_housing', '08_housing_reasons']:
    print(f'=== {sheet} ===')
    ws = wb[sheet]
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if any(c is not None for c in row):
            vals = [str(c)[:30] if c is not None else '' for c in row[:7]]
            print(f'  row {i}: {vals}')
    print()
