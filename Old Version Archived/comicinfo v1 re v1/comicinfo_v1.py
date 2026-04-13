import os
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

# ===== 設定 =====
EXCEL_FILE = "input_v1.xlsx"
OUTPUT_DIR = "output"

# ===== 工具 =====
def sanitize_folder_name(name):
    invalid_chars = r'<>:"/\\|?*'
    for c in invalid_chars:
        name = name.replace(c, "_")
    return name.strip()

def add_tag(parent, tag, text):
    elem = ET.SubElement(parent, tag)
    elem.text = "" if pd.isna(text) else str(text)

def prettify(elem):
    rough = ET.tostring(elem, encoding="utf-8")
    reparsed = minidom.parseString(rough)

    # 遍歷所有元素，空文字也保留
    for node in reparsed.getElementsByTagName("*"):
        if node.firstChild is None:
            node.appendChild(reparsed.createTextNode(""))

    return reparsed.toprettyxml(indent="    ", encoding="utf-8")

def to_int_str(value):
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except:
        return str(value)

def create_xml(row):
    root = ET.Element(
        "ComicInfo",
        {
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "xsi:noNamespaceSchemaLocation": "https://raw.githubusercontent.com/anansi-project/comicinfo/main/schema/v2.0/ComicInfo.xsd"
        }
    )

    # ===== 基本欄位 =====
    add_tag(root, "Title", row.get("Title"))
    add_tag(root, "Series", row.get("Series"))
    add_tag(root, "Number", to_int_str(row.get("Number")))
    add_tag(root, "Count", to_int_str(row.get("Count")))

    summary = ET.SubElement(root, "Summary")
    summary.text = "" if pd.isna(row.get("Summary")) else str(row.get("Summary"))

    add_tag(root, "Year", to_int_str(row.get("Year")))
    add_tag(root, "Month", to_int_str(row.get("Month")))
    add_tag(root, "Day", to_int_str(row.get("Day")))
    add_tag(root, "Writer", row.get("Writer"))

    # 固定空欄
    for tag in ["Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "Imprint"]:
        add_tag(root, tag, "")

    add_tag(root, "Publisher", row.get("Publisher"))
    add_tag(root, "Genre", row.get("Genre"))
    add_tag(root, "Web", row.get("Web"))
    add_tag(root, "PageCount", to_int_str(row.get("PageCount")))

    # 固定值
    add_tag(root, "LanguageISO", "zh-tw")
    add_tag(root, "Format", "Digital")
    add_tag(root, "BlackAndWhite", "No")
    add_tag(root, "Manga", "YesAndRightToLeft")

    return root

# ===== 主流程 =====
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(EXCEL_FILE)

    for idx, row in df.iterrows():
        title = row.get("Title")

        if pd.isna(title):
            print(f"⚠️ 第 {idx+1} 列沒有 Title，跳過")
            continue

        folder_name = sanitize_folder_name(str(title))
        target_dir = os.path.join(OUTPUT_DIR, folder_name)
        os.makedirs(target_dir, exist_ok=True)

        xml_root = create_xml(row)
        xml_data = prettify(xml_root)

        output_file = os.path.join(target_dir, "comicinfo.xml")

        with open(output_file, "wb") as f:
            f.write(xml_data)

        print(f"✔ 已生成: {output_file}")

if __name__ == "__main__":
    main()