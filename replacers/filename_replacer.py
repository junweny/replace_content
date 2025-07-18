import os
import re

def replace_filename(file_path, replacements):
    dir_name, base_name = os.path.split(file_path)
    new_name = base_name
    for old, new in replacements:
        new_name = new_name.replace(old, new)
    # 替换新文件名中的特殊字符（如！等）为下划线
    new_name = re.sub(r'[！!@#$%^&*?<>|:"\'/]', '_', new_name)
    if new_name != base_name:
        new_path = os.path.join(dir_name, new_name)
        os.rename(file_path, new_path)
        return new_path
    return file_path 