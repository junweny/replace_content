def to_fullwidth(s):
    return ''.join(chr(ord(c) + 0xFEE0) if 0x21 <= ord(c) <= 0x7E else c for c in s)

def to_halfwidth(s):
    return ''.join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c for c in s) 