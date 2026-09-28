"""Debug helper: print the first rows of an .xlsx file without third-party libraries.

Usage:
    uv run read_xlsx.py "Scan Ex.xlsx"
    uv run read_xlsx.py "Scan Ex.xlsx" --sheet 2 --rows 50
"""
import argparse
import os
import zipfile
import xml.etree.ElementTree as ET

NS = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def load_shared_strings(zip_ref):
    try:
        with zip_ref.open('xl/sharedStrings.xml') as f:
            root = ET.parse(f).getroot()
    except KeyError:
        return []

    # One <si> per string; rich text splits a single string across several <t> elements
    return [''.join(t.text or '' for t in si.iter(f"{{{NS['ns']}}}t"))
            for si in root.findall('ns:si', NS)]


def read_cell_value(c_elem, shared_strings):
    cell_type = c_elem.get('t')
    if cell_type == 'inlineStr':
        return ''.join(t.text or '' for t in c_elem.iter(f"{{{NS['ns']}}}t"))

    v_elem = c_elem.find('ns:v', NS)
    if v_elem is None:
        return None
    if cell_type == 's':
        return shared_strings[int(v_elem.text)]
    return v_elem.text


def parse_xlsx(file_path, sheet_number=1, max_rows=20):
    print(f"Reading file: {file_path}")
    if not os.path.exists(file_path):
        print("File does not exist")
        return

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        shared_strings = load_shared_strings(zip_ref)
        sheet_path = f'xl/worksheets/sheet{sheet_number}.xml'

        try:
            with zip_ref.open(sheet_path) as f:
                root = ET.parse(f).getroot()
        except KeyError:
            print(f"Sheet not found: {sheet_path}")
            return

        rows = root.findall('.//ns:row', NS)
        for row_elem in rows[:max_rows]:
            row_data = {}
            for c_elem in row_elem.findall('ns:c', NS):
                col = ''.join(ch for ch in c_elem.get('r', '') if ch.isalpha())
                row_data[col] = read_cell_value(c_elem, shared_strings)
            print(f"Row {row_elem.get('r')}: {row_data}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Print the first rows of an .xlsx sheet.')
    parser.add_argument('file', help='path to the .xlsx file')
    parser.add_argument('--sheet', type=int, default=1, help='sheet number (default: 1)')
    parser.add_argument('--rows', type=int, default=20, help='number of rows to print (default: 20)')
    args = parser.parse_args()
    parse_xlsx(args.file, args.sheet, args.rows)
