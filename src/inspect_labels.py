import json

def list_sheet_labels_with_cols(sheet_name):
    try:
        with open('data/dsf_inventory.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        return
    
    sheet = data['sheets'].get(sheet_name)
    if not sheet:
        return

    print(f"--- Layout for {sheet_name} ---")
    
    # 1. Show Input Columns
    input_cols = sheet.get('input_columns', [])
    for ic in input_cols:
        print(f"INPUT_COL: {ic.get('letter')} ({ic.get('header_value')[:30]}...)")
        
    # 2. Show Labels with their position
    seen = set()
    fields = sheet.get('fields', [])
    # Sort by row
    fields.sort(key=lambda x: x['row'])
    
    for field in fields:
        label = field.get('label')
        if label and len(label) > 3:
             if label not in seen:
                 print(f"LABEL: Row {field.get('row')} Col {field.get('column_letter')} -> '{label}'")
                 seen.add(label)

if __name__ == "__main__":
    list_sheet_labels_with_cols("NOTE 4 ")
    list_sheet_labels_with_cols("NOTE 13 ")
