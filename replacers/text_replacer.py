import re
import codecs
from utils.encoding_utils import detect_encoding

def replace_in_text_file(file_path, replacements, full_word=False):
    encoding = detect_encoding(file_path)
    with codecs.open(file_path, 'r', encoding=encoding) as f:
        content = f.read()
    for old, new in replacements:
        if full_word:
            content = re.sub(rf'\b{re.escape(old)}\b', new, content)
        else:
            content = content.replace(old, new)
    with codecs.open(file_path, 'w', encoding=encoding) as f:
        f.write(content) 