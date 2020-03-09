
def sanitize(s: str):
    s_sanitized = s.strip().split(' ')
    s_sanitized = ''.join(s_sanitized).split('_')
    s_sanitized = ''.join(s_sanitized)
    s_sanitized = ''.join(e for e in s_sanitized if e.isalnum())
    return s_sanitized.lower()


def merge_lists(l1: list, l2: list) -> list:
    difference = list(set(l2) - set(l1))
    return list(l1) + difference


class safelist(list):
    def get(self, index, default=None):
        try:
            return self.__getitem__(index)
        except IndexError:
            return default


# Just a few colors to use in the console logs.
ENDC = '\033[0m'
RED = '\033[91m'
OKGREEN = '\033[92m'
YELLOW = '\033[93m'
