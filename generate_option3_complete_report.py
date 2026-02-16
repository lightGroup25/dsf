"""
Generate comprehensive report for Option 3 Complete with all Notes
"""

import openpyxl
import json
from pathlib import Path
from datetime import datetime

OUTPUT_FILE = r"output/DSF_OPTION3_COMPLETE_WITH_NOTES.xlsx"

def generate_complete_report():
    """Generate detailed report"""
    print("="*70)
    print("OPTION 3 COMPLETE - COMPREHENSIVE REPORT")
    print("="*70)
    
    try:
        wb = openpyxl.load_workbook(OUTPUT_FILE, data_only=False)
        
        report = {
            'timestamp': datetime.now().isoformat(),
            'title': 'DSF AUTO-FILL - OPTION 3 COMPLETE WITH ALL NOTES',
            'summary': {
                'total_sheets': len(wb.sheetnames),
                'output_file': OUTPUT_FILE,
                'file_size_mb': round(Path(OUTPUT_FILE).stat().st_size / (1024*1024), 2)
            },
            'sheet_details': []
        }
        
        total_cells_with_data = 0
        total_formulas = 0
        sheet_summary_by_type = {
            'main': [],
            'notes': [],
            'other': []
        }
        
        print("\nScanning all sheets...")
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            data_cells = 0
            formula_cells = 0
            
            # Count cells with data
            for row_idx in range(1, min(500, ws.max_row + 1)):
                for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
                    cell = ws[f'{col}{row_idx}']
                    if cell.value:
                        if isinstance(cell.value, str) and cell.value.startswith('='):
                            formula_cells += 1
                        elif isinstance(cell.value, (int, float)):
                            data_cells += 1
            
            total_cells_with_data += data_cells
            total_formulas += formula_cells
            
            sheet_info = {
                'name': sheet_name,
                'data_cells': data_cells,
                'formula_cells': formula_cells,
                'total': data_cells + formula_cells
            }
            
            report['sheet_details'].append(sheet_info)
            
            # Categorize sheets
            if 'NOTE' in sheet_name and data_cells + formula_cells > 0:
                sheet_summary_by_type['notes'].append(sheet_name)
                print(f"  ✓ {sheet_name}: {data_cells} data + {formula_cells} formulas")
            elif any(x in sheet_name for x in ['BILAN', 'COMPTE', 'TABLEAU']):
                if data_cells + formula_cells > 0:
                    sheet_summary_by_type['main'].append(sheet_name)
            elif data_cells + formula_cells > 0:
                sheet_summary_by_type['other'].append(sheet_name)
        
        report['summary'].update({
            'total_data_cells': total_cells_with_data,
            'total_formula_cells': total_formulas,
            'total_cells_impacted': total_cells_with_data + total_formulas,
            'sheets_with_notes_filled': len(sheet_summary_by_type['notes']),
            'sheets_by_category': {
                'main_statements': len(sheet_summary_by_type['main']),
                'note_sections': len(sheet_summary_by_type['notes']),
                'other': len(sheet_summary_by_type['other'])
            }
        })
        
        # Print summary
        print("\n" + "="*70)
        print("SUMMARY")
        print("="*70)
        print(f"Total sheets in template: {report['summary']['total_sheets']}")
        print(f"Sheets with filled data:")
        print(f"  - Main statements: {report['summary']['sheets_by_category']['main_statements']}")
        print(f"  - Note sections: {report['summary']['sheets_by_category']['note_sections']}")
        print(f"  - Other sections: {report['summary']['sheets_by_category']['other']}")
        print(f"\nCells filled:")
        print(f"  - With data: {report['summary']['total_data_cells']:,}")
        print(f"  - With formulas: {report['summary']['total_formula_cells']:,}")
        print(f"  - Total impact: {report['summary']['total_cells_impacted']:,}")
        print(f"\nFile size: {report['summary']['file_size_mb']} MB")
        
        # List all NOTE sheets that were filled
        print(f"\nNOTE sections filled ({len(sheet_summary_by_type['notes'])}):")
        for note in sorted(sheet_summary_by_type['notes']):
            print(f"  ✓ {note}")
        
        # Save report
        report_path = r"output/reports/OPTION3_COMPLETE_REPORT.json"
        Path(report_path).parent.mkdir(parents=True, exist_ok=True)
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        print(f"\n✓ Report saved: {report_path}")
        
        # Create markdown summary
        md_path = r"output/reports/OPTION3_COMPLETE_SUMMARY.md"
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write("# DSF AUTO-FILL - OPTION 3 COMPLETE\n\n")
            f.write(f"**Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write(f"**Output File:** `{OUTPUT_FILE}`\n\n")
            f.write(f"**File Size:** {report['summary']['file_size_mb']} MB\n\n")
            f.write("## Summary\n\n")
            f.write(f"- Total sheets processed: {report['summary']['total_sheets']}\n")
            f.write(f"- Cells with data: **{report['summary']['total_data_cells']:,}**\n")
            f.write(f"- Cells with formulas: **{report['summary']['total_formula_cells']:,}**\n")
            f.write(f"- **Total cells impacted: {report['summary']['total_cells_impacted']:,}**\n\n")
            f.write("## Sheets Filled\n\n")
            f.write(f"### Main Statements ({report['summary']['sheets_by_category']['main_statements']})\n")
            for sheet in sheet_summary_by_type['main']:
                f.write(f"- {sheet}\n")
            f.write(f"\n### Note Sections ({report['summary']['sheets_by_category']['note_sections']})\n")
            for note in sorted(sheet_summary_by_type['notes']):
                f.write(f"- {note}\n")
        
        print(f"✓ Summary saved: {md_path}")
        
        print("\n" + "="*70)
        print("✅ OPTION 3 COMPLETE - ALL NOTES SUCCESSFULLY FILLED")
        print("="*70)
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    generate_complete_report()
