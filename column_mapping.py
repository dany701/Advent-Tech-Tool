def generate_column_mapping():
    mapping = {}
    for i in range(1, 16385):
        col_letter = ""
        n = i
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            col_letter = chr(65 + remainder) + col_letter
        mapping[col_letter] = i - 1  # 0-based index
    return mapping
