import zipfile
import xml.etree.ElementTree as ET
import os
import re
import pandas as pd
from tqdm import tqdm


import sys

xlsx_path = "sample.xlsx"

class CONVERT_TO_XML:
    def __init__(self, xlsx_path):
        self.xlsx_path = xlsx_path
        self.output_folder = "data/sample_unzipped"

    def unzip_xlsx_to_xml_folder(self):
        os.makedirs(self.output_folder, exist_ok=True)
        with zipfile.ZipFile(self.xlsx_path, 'r') as zip_ref:
            zip_ref.extractall(self.output_folder)


class INFO_XML(CONVERT_TO_XML):
    def __init__(self, xlsx_path):
        super().__init__(xlsx_path)
        self.shared_strings_path = os.path.join(self.output_folder, "xl/sharedStrings.xml")
        self.set_sheetname()

    def init_parameter(self):
        self.ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # Parse shared strings and sheet XML once for reuse
        self.ss_tree = ET.parse(self.shared_strings_path)
        self.ss_root = self.ss_tree.getroot()
        self.si_list = self.ss_root.findall("a:si", self.ns)

        self.sheet_tree = ET.parse(self.sheet_xml_path)
        self.sheet_root = self.sheet_tree.getroot()

    def set_sheetname(self, sheet_name="sheet1"):
        self.sheet_xml_path = os.path.join(self.output_folder, f"xl/worksheets/{sheet_name}.xml")

    def get_shared_string_index(self, cell_ref):
        cell = self.sheet_root.find(f".//a:c[@r='{cell_ref}']", self.ns)
        if cell is None:
            raise ValueError("Không tìm thấy")
        return int(cell.find("a:v", self.ns).text)

    def get_rich_text_info(self, cell_ref):
        try:
            index = self.get_shared_string_index(cell_ref)
        except:
            return None

        if index >= len(self.si_list):
            return None

        si = self.si_list[index]
        sentence = []

        for r in si.findall("a:r", self.ns):
            t_elem = r.find("a:t", self.ns)
            text = t_elem.text if t_elem is not None else ""

            rPr = r.find("a:rPr", self.ns)
            font = (
                rPr.find("a:rFont", self.ns).attrib.get("val")
                if rPr is not None and rPr.find("a:rFont", self.ns) is not None
                else "default"
            )
            color = (
                rPr.find("a:color", self.ns).attrib.get("rgb")
                if rPr is not None and rPr.find("a:color", self.ns) is not None
                else "default"
            )

            for char in text:
                sentence.append({"Text": char, "Font": font, "Color": color})

        return sentence if sentence else None


    def count_rows_in_column(self, column_letter):
        count = 0
        for cell in self.sheet_root.findall(".//a:sheetData/a:row/a:c", self.ns):
            ref = cell.attrib.get("r", "")
            if ref.startswith(column_letter) and cell.find("a:v", self.ns) is not None:
                count += 1
        return count

    def get_column_headers(self):
        headers = {}
        for cell in self.sheet_root.findall(".//a:sheetData/a:row[@r='1']/a:c", self.ns):
            ref = cell.attrib.get("r", "")
            col_letter = re.match(r"[A-Z]+", ref).group()
            v = cell.find("a:v", self.ns)
            if v is not None:
                try:
                    idx = int(v.text)
                    si = self.si_list[idx]
                    t = si.find("a:t", self.ns)
                    if t is not None:
                        headers[t.text.strip()] = col_letter
                except:
                    continue
        return headers

    def count_columns(self):
        row = self.sheet_root.find(".//a:sheetData/a:row[@r='1']", self.ns)
        return len(row.findall("a:c", self.ns)) if row is not None else 0


class PROCESS_XLXS(INFO_XML):
    def __init__(self, xlsx_path):
        super().__init__(xlsx_path)

    def process(self, name_column):
        self.unzip_xlsx_to_xml_folder()
        self.init_parameter()
        column_map = self.get_column_headers()
        if name_column not in column_map:
            return None
        column = column_map[name_column]
        num_row = self.count_rows_in_column(column)
        data = {"Name": [], "Config": []}
        for idx in tqdm(range(2, num_row + 2), desc="Process file", unit="lines"):
            name = f"{column}{idx}"# C2
            sequence = self.get_rich_text_info(name)
            data["Name"].append(name)
            data["Config"].append(sequence)
        return pd.DataFrame(data)##
