import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw = f.read(4096)
    return chardet.detect(raw)['encoding'] or 'utf-8' 