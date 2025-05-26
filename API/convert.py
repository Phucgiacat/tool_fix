import zipfile
import xml.etree.ElementTree as ET
import os
import re
from tqdm import tqdm
import pandas as pd

xlsx_path = "sample.xlsx"
class CONVERT_TO_XML:
    def __init__(self, xlsx_path):
        self. xlsx_path = xlsx_path
        self.output_folder = "sample_unzipped"

    def unzip_xlsx_to_xml_folder(self):
        os.makedirs(self.output_folder, exist_ok=True)   
        with zipfile.ZipFile(self.xlsx_path, 'r') as zip_ref:
            zip_ref.extractall(self.output_folder)
        print(f"Đã giải nén {xlsx_path} vào thư mục {self.output_folder}")



class INFO_XML(CONVERT_TO_XML):
    def __init__(self, xlsx_path):
        super().__init__(xlsx_path)
        self.shared_strings_path = os.path.join(self.output_folder, "xl/sharedStrings.xml")
        self.set_sheetname()

    def set_sheetname(self, sheet_name="sheet1"):
        self.sheet_xml_path = self.sheet_xml_path = os.path.join(self.output_folder, f"xl/worksheets/{sheet_name}.xml")
    
    def get_shared_string_index(self, cell_ref):
        sheet_xml_path = self.sheet_xml_path
        tree = ET.parse(sheet_xml_path)
        root = tree.getroot()
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        cell = root.find(f".//a:c[@r='{cell_ref}']", ns)
        if cell is None:
            raise ValueError("Không tìm thấy")
        return int(cell.find("a:v", ns).text)


    def get_rich_text_info( self,cell_ref):
        index = self.get_shared_string_index(cell_ref)
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        tree = ET.parse(self.shared_strings_path)
        root = tree.getroot()

        si_list = root.findall("a:si", ns)
        if index >= len(si_list):
            print("Index không tồn tại trong sharedStrings")
            return

        si = si_list[index]
        sentence = []
        for r in si.findall("a:r", ns):
            text = r.find("a:t", ns).text or ""
            rPr = r.find("a:rPr", ns)

            font = rPr.find("a:rFont", ns).attrib.get("val") if rPr is not None and rPr.find("a:rFont", ns) is not None else "default"
            color = rPr.find("a:color", ns).attrib.get("rgb") if rPr is not None and rPr.find("a:color", ns) is not None else "default"

            word = {"Text": text , "Font": {font} , "Color": color}
            sentence.append(word)
        return sentence
    
    ## Hàm Đếm số lượng 
    def count_rows_in_column(self, column_letter):
        sheet_xml_path = self.sheet_xml_path
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        tree = ET.parse(sheet_xml_path)
        root = tree.getroot()
        count = 0
        for cell in root.findall(".//a:sheetData/a:row/a:c", ns):
            ref = cell.attrib.get("r", "")
            if ref.startswith(column_letter):
                # Đảm bảo ô này thực sự có dữ liệu
                if cell.find("a:v", ns) is not None:
                    count += 1
        return count
    
    def get_column_headers(self):
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # 1. Parse shared strings
        ss_tree = ET.parse(self.shared_strings_path)
        ss_root = ss_tree.getroot()
        si_list = ss_root.findall("a:si", ns)

        # 2. Parse sheet1.xml
        sheet_tree = ET.parse(self.sheet_xml_path)
        sheet_root = sheet_tree.getroot()

        headers = {}

        # Tìm dòng đầu tiên (header)
        for cell in sheet_root.findall(".//a:sheetData/a:row[@r='1']/a:c", ns):
            ref = cell.attrib.get("r")  # Ví dụ "D1"
            col_letter = re.match(r"[A-Z]+", ref).group()

            # Lấy index từ shared string
            v = cell.find("a:v", ns)
            if v is not None:
                idx = int(v.text)
                si = si_list[idx]
                t = si.find("a:t", ns)
                if t is not None:
                    headers[t.text.strip()] = col_letter
        return headers
    
    def count_columns(self):
        ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        tree = ET.parse(self.sheet_xml_path)
        root = tree.getroot()

        # Tìm dòng đầu tiên (thường chứa tiêu đề)
        row = root.find(".//a:sheetData/a:row[@r='1']", ns)
        if row is None:
            print("❌ Không tìm thấy dòng tiêu đề (row 1).")
            return 0

        column_count = len(row.findall("a:c", ns))
        print(f"📊 Sheet có {column_count} cột.")
        return column_count
    
class PROCESS_XLXS(INFO_XML):
    def __init__(self, xlsx_path):
        super().__init__(xlsx_path)

    def process(self, name_column):
        if os.path.exists(self.output_folder) == False:
            self.unzip_xlsx_to_xml_folder()
        column = self.get_column_headers()
        if name_column not in column.keys():
            return None
        column = column[name_column]
        num_row = self.count_rows_in_column(column)
        dict_ = {
            "Name": [],
            "Config": []
        }
        for idx in tqdm(range(2,num_row), desc="Procss file:", unit="lines"):
            name = f"{column}{idx}"
            sequence = self.get_rich_text_info(name)
            dict_["Name"].append(name)
            dict_["Config"].append(sequence)
        
        return pd.DataFrame(dict_)