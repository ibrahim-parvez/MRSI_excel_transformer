def normalize_name(s):
    if s is None:
        return ''
    return ' '.join(str(s).split()).lower()
