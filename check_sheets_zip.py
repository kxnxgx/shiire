import zipfile
import xml.etree.ElementTree as ET

try:
    with zipfile.ZipFile('シレトコ分析.xlsx', 'r') as z:
        # Read workbook.xml
        wb_xml = z.read('xl/workbook.xml')
        root = ET.fromstring(wb_xml)
        
        # namespaces
        ns = {'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
              
        print("Sheets in workbook.xml:")
        for sheet in root.findall('.//main:sheet', ns):
            name = sheet.attrib.get('name')
            state = sheet.attrib.get('state', 'visible')
            print(f"- {name}: state={state}")
            
except Exception as e:
    print("Error:", e)
