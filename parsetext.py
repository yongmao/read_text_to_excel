# 以分隔符将前后两部分分隔成 key value
def parse_to_pair(line, delimiters=None):
    if delimiters is None:
        delimiters = [':', '：']

    cells = split_string_with_multiple_delimiters(cell_line, delimiters)
    # print(result)
    cell_name = cells[0].strip()
