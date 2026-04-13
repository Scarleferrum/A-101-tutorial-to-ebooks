import os
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

# ===== 設定 =====
EXCEL_FILE = "input_v1.5.1.xlsx"
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

def add_ISO(parent, tag, text):
    elem = ET.SubElement(parent, tag)
    elem.text = "zh-tw" if pd.isna(text) else str(text)

def add_For(parent, tag, text):
    elem = ET.SubElement(parent, tag)
    elem.text = "Digital" if pd.isna(text) else str(text)

def add_BNW(parent, tag, text):
    elem = ET.SubElement(parent, tag)
    elem.text = "No" if pd.isna(text) else str(text)

def add_RTL(parent, tag, text):
    elem = ET.SubElement(parent, tag)
    elem.text = "YesAndRightToLeft" if pd.isna(text) else str(text)

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

    add_tag(root, "Summary", row.get("Summary"))

    for tag in ["Year", "Month", "Day"]:
        add_tag(root, tag, to_int_str(row.get(tag)))
    
    add_tag(root, "Writer", row.get("Writer"))

    for tag in ["Penciller", "Inker", "Colorist", "Letterer", "CoverArtist", "Editor", "Imprint"]:
        add_tag(root, tag, "")

    # For loops
    for tag in ["Publisher", "Genre", "Tags", "Web"]:
        add_tag(root, tag, row.get(tag))

    add_tag(root, "PageCount", "")
    add_ISO(root, "LanguageISO", row.get("LanguageISO"))
    add_For(root, "Format", row.get("Format"))
    add_BNW(root, "BlackAndWhite", row.get("BlackAndWhite"))
    add_RTL(root, "Manga", row.get("Manga"))
    add_tag(root, "Characters",row.get("Characters"))
    add_tag(root, "Teams", row.get("Teams"))

    # 固定值
    

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