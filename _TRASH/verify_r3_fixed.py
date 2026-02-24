
import openpyxl

def verify_r3_fixed(filename):
    print(f"Verifying Fiche R3 in {filename}...\n")
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
        ws = wb["Fiche R3"]
        
        checks = [
            # Headers
            ("A4", "GULFCAM SAS"),
            ("A5", "M050900027774W"),
            
            # Dirigeants (rows 11-13)
            ("A11", "NYODOG"),
            ("B11", "PERRIAL JEAN"),
            ("A12", "BOU"),
            ("A13", "MIKE"),
            
            # Conseil - First section (rows 26-29)
            ("B26", "MBAYEN"),
            ("C26", "RENE"),
            ("D26", "PRESIDENT CONSEIL SURVEILLANCE"),
            
            ("B27", "NYODOG"),
            ("C27", "PERRIAL JEAN"),
            ("D27", "PRESIDENT DU DIRECTOIRE"),
            
            ("B28", "TCHOUTA MOUSSA"),
            ("C28", "ESTHER"),
            ("D28", "MEMBRE"),
            
            ("B29", "NJOVAGE"),
            ("C29", "GEORGES"),
            
            # Conseil - Second section (rows 31-35, with combined names)
            ("B31", "MBAYEN HEGBA FLORIAN"),
            ("E31", "DOUALA, CAMEROUN"),
            
            ("B33", "NGO MBAYEN Epse MALANANGE DORA"),
            ("E33", "DOUALA, CAMEROUN"),
            
            ("B34", "FAYCAL ABDOULAYE"),
            ("E34", "YAOUNDE, CAMEROUN"),
            
            ("B35", "WOLINO Epse BOOTO A N COLETTE"),
            ("E35", "DOUALA, CAMEROUN"),
        ]
        
        all_passed = True
        for coord, expected in checks:
            val = ws[coord].value
            if val is None:
                val = ""
            passed = expected.upper() in str(val).upper()
            status = "✓" if passed else "✗"
            print(f"[{status}] {coord}: {val}")
            if not passed:
                all_passed = False
                print(f"         Expected substring: {expected}")
        
        if all_passed:
            print("\n✓ SUCCESS: Fiche R3 is correctly filled!")
        else:
            print("\n✗ FAILURE: Some checks failed")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    verify_r3_fixed(r"c:\Users\Emmaneul Ambadiang\Desktop\DSF\output\DSF_OUTPUT_2024_20260213_190157.xlsx")
