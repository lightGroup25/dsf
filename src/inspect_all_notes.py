import json

def list_all_note_titles():
    try:
        with open('data/dsf_inventory.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        return
    
    sheets = data['sheets']
    for name, sheet in sheets.items():
        if "NOTE" in name.upper():
            # Find the title (usually row 5 or so)
            title = ""
            for field in sheet.get('fields', []):
                 if field['row'] < 8 and field['column'] == 2: # Column B usually has the title
                     if field.get('label') and len(field['label']) > 5:
                         title = field['label'].replace('\n', ' ')
                         break
            print(f"{name}: {title}")

if __name__ == "__main__":
    list_all_note_titles()
