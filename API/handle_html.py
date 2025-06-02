from bs4 import BeautifulSoup
import re
from langdetect import detect, LangDetectException
import xlsxwriter


def extract_text_with_color(html_path):
    # Map tên màu sang mã hex
    color_name_to_hex = {
        'black': '#000000',
        'red': '#FF0000',
        'blue': '#0000FF',
        'green': '#00AA00',   # hoặc '#008000' tùy hệ thống
    }

    with open(html_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    table = soup.find("table")
    results = []

    for row_idx, tr in enumerate(table.find_all("tr")):
        row_data = []
        for td in tr.find_all(["td", "th"]):
            cell_chunks = []
            spans = td.find_all("span")
            if spans:
                for span in spans:
                    text = span.get_text(strip=True)
                    style = span.get("style", "").lower()
                    match = re.search(r"color:\s*([^;]+)", style)
                    if match:
                        color_code = match.group(1).strip()
                        if color_code.startswith("rgb"):
                            r, g, b = map(int, color_code.replace("rgb(", "").replace(")", "").split(","))
                            hex_color = "#{:02X}{:02X}{:02X}".format(r, g, b)
                        elif color_code.startswith("#"):
                            hex_color = color_code
                        else:
                            hex_color = color_name_to_hex.get(color_code, "#000000")
                    else:
                        hex_color = "#000000"
                    cell_chunks.append((text, hex_color))
            else:
                # Nếu không có span, lấy toàn bộ text và color từ td
                text = td.get_text(strip=True)
                style = td.get("style", "").lower()
                match = re.search(r"color:\s*([^;]+)", style)
                if match:
                    color_code = match.group(1).strip()
                    if color_code.startswith("rgb"):
                        r, g, b = map(int, color_code.replace("rgb(", "").replace(")", "").split(","))
                        hex_color = "#{:02X}{:02X}{:02X}".format(r, g, b)
                    elif color_code.startswith("#"):
                        hex_color = color_code
                    else:
                        hex_color = color_name_to_hex.get(color_code, "#000000")
                else:
                    hex_color = "#000000"
                cell_chunks.append((text, hex_color))

            row_data.append(cell_chunks)
        results.append(row_data)
    return results
def write_colored_excel_from_chunks(data, output_path):
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()

    # Định nghĩa 4 màu bạn yêu cầu
    red   = workbook.add_format({'font_color': 'red'})
    blue  = workbook.add_format({'font_color': 'blue'})
    green = workbook.add_format({'font_color': 'green'})
    black = workbook.add_format({'font_color': 'black'})
    header = workbook.add_format({'bold': True, 'align': 'center'})

    # Tạo bộ map hex → format
    color_map = {
        '#FF0000': red,
        '#0000FF': blue,
        '#00AA00': green,
        '#008000': green,
        '#000000': black
    }

    # Ghi dữ liệu ra Excel
    for row_idx, row in enumerate(data):
        for col_idx, cell in enumerate(row):
            if row_idx == 0:
                # Dòng đầu coi là tiêu đề
                text = "".join([chunk[0] for chunk in cell if chunk[0].strip()])
                worksheet.write(row_idx, col_idx, text, header)
                continue

            # Loại bỏ các đoạn text rỗng
            non_empty_chunks = [(t, c) for t, c in cell if t.strip()]
            

            if not non_empty_chunks:
                worksheet.write(row_idx, col_idx, "")
            elif len(non_empty_chunks) == 1:
                text, hex_color = non_empty_chunks[0]
                fmt = color_map.get(hex_color.upper(), black)
                worksheet.write(row_idx, col_idx, text, fmt)
            else:
                rich_text = []
                seqence = " ".join([text for text, _ in non_empty_chunks]).strip()
                
                # ✅ Check trước khi detect ngôn ngữ
                try:
                    add_space = len(seqence) > 3 and detect(seqence) == "vi"
                except LangDetectException:
                    add_space = False

                for i, (text, hex_color) in enumerate(non_empty_chunks):
                    fmt = color_map.get(hex_color.upper(), black)
                    if add_space and i < len(non_empty_chunks) - 1:
                        rich_text.append(fmt)
                        rich_text.append(text + " ")
                    else:
                        rich_text.append(fmt)
                        rich_text.append(text)
                worksheet.write_rich_string(row_idx, col_idx, *rich_text)
    worksheet.set_column('A:Z', 40)  # Tùy chỉnh độ rộng nếu cần
    workbook.close()
    print("✅ Đã lưu Excel ra:", output_path)