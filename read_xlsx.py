import zipfile
import xml.etree.ElementTree as ET
import os

def parse_xlsx(file_path):
    print(f"Reading file: {file_path}")
    if not os.path.exists(file_path):
        print("File does not exist")
        return

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        # Load shared strings
        shared_strings = []
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                # Namespaces
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for t in root.findall('.//ns:t', ns):
                    shared_strings.append(t.text)
        except KeyError:
            print("No shared strings found")

        # Load sheet1
        try:
            with zip_ref.open('xl/worksheets/sheet1.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                rows = []
                for row_elem in root.findall('.//ns:row', ns):
                    row_data = {}
                    for c_elem in row_elem.findall('ns:c', ns):
                        r = c_elem.get('r') # e.g. A1
                        # Parse col index from r
                        col_str = ''.join([char for char in r if char.isalpha()])
                        t = c_elem.get('t')
                        v_elem = c_elem.find('ns:v', ns)
                        val = None
                        if v_elem is not None:
                            val = v_elem.text
                            if t == 's':
                                val = shared_strings[int(val)]
                        row_data[col_str] = val
                    rows.append(row_data)
                
                # Print first 20 rows
                for idx, row in enumerate(rows[:20]):
                    print(f"Row {idx+1}: {row}")
        except Exception as e:
            print(f"Error reading sheet1: {e}")

parse_xlsx("Scan Ex.xlsx")
