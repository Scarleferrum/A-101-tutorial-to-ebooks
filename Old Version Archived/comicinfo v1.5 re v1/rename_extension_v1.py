import os

def change_file_extension(folder_path, old_ext, new_ext):
    """
    批量更改指定資料夾下檔案的副檔名
    :param folder_path: 資料夾完整路徑
    :param old_ext: 原副檔名 (例如 '.zip')
    :param new_ext: 新副檔名 (例如 '.cbz')
    """
    if not os.path.isdir(folder_path):
        print(f"資料夾不存在: {folder_path}")
        return

    # 確保副檔名以點開頭
    if not old_ext.startswith('.'):
        old_ext = '.' + old_ext
    if not new_ext.startswith('.'):
        new_ext = '.' + new_ext

    # 遍歷資料夾內所有檔案
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(old_ext.lower()):
            old_file = os.path.join(folder_path, filename)
            new_filename = filename[:-len(old_ext)] + new_ext
            new_file = os.path.join(folder_path, new_filename)
            os.rename(old_file, new_file)
            print(f"已更改: {filename} → {new_filename}")

if __name__ == "__main__":
    # 取得腳本所在資料夾
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_dir, "output")

    # 更改 output 資料夾內所有 .zip 為 .cbz
    change_file_extension(output_folder, ".zip", ".cbz")