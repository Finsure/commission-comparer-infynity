
def merge_lists(l1: list, l2: list) -> list:
    difference = list(set(l2) - set(l1))
    return list(l1) + difference


class safelist(list):
    def get(self, index, default=None):
        try:
            return self.__getitem__(index)
        except IndexError:
            return default
