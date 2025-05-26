import zipfile
import xml.etree.ElementTree as ET

xlsx_path = "sample.xlsx"

# === Bước 1: Lấy chỉ số shared string từ cell D4 trong sheet1.xml ===
with zipfile.ZipFile(xlsx_path, 'r') as z:
    with z.open('xl/worksheets/sheet1.xml') as f:
        sheet_tree = ET.parse(f)
        sheet_root = sheet_tree.getroot()
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # Tìm ô D4
        cell = sheet_root.find(".//a:c[@r='D6']", ns)
        if cell is None:
            print("❌ Không tìm thấy ô D4.")
            exit()
        shared_string_index = int(cell.find("a:v", ns).text)

    # === Bước 2: Lấy chuỗi định dạng từ sharedStrings.xml ===
    with z.open('xl/sharedStrings.xml') as f:
        ss_tree = ET.parse(f)
        ss_root = ss_tree.getroot()

        # Lấy phần tử <si> tương ứng chỉ số
        si = ss_root.findall("a:si", ns)[shared_string_index]

        print(f"🎯 Nội dung ô D1:")

        for r in si.findall("a:r", ns):
            text = r.find("a:t", ns).text
            rPr = r.find("a:rPr", ns)
            
            # Lấy màu nếu có
            color_tag = rPr.find("a:color", ns) if rPr is not None else None
            color = color_tag.attrib.get("rgb") if color_tag is not None else "default"
            
            # Lấy font nếu có
            font_tag = rPr.find("a:rFont", ns) if rPr is not None else None
            font = font_tag.attrib.get("val") if font_tag is not None else "default"
            print(f"  - Text: {text} | Font: {font} | Color: {color}")
