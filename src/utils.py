def merge_lists(l1: list, l2: list) -> list:
    difference = list(set(l2) - set(l1))
    return list(l1) + difference
