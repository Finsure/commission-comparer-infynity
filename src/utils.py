
def sanitize(s: str):
    s_sanitized = s.strip().split(' ')
    s_sanitized = ''.join(s_sanitized).split('_')
    s_sanitized = ''.join(s_sanitized)
    s_sanitized = ''.join(e for e in s_sanitized if e.isalnum())
    return s_sanitized.lower()


def merge_lists(l1: list, l2: list) -> list:
    difference = list(set(l2) - set(l1))
    return list(l1) + difference


def money_to_float(value: str):
    new_value = value
    new_value = new_value.replace('$', '')
    new_value = new_value.replace(',', '')
    return float(new_value)


def compare_numbers(n1, n2, margin):
    n1val = str(n1)
    n2val = str(n2)

    try:
        n1val = money_to_float(n1val)
        n2val = money_to_float(n2val)
    except ValueError:
        if n1val == '' or n2val == '':
            return n1val == n2val
        return False

    return abs(n1val - n2val) <= margin + 0.000001


class safelist(list):
    def get(self, index, default=None):
        try:
            return self.__getitem__(index)
        except IndexError:
            return default


# Just a few colors to use in the console logs.
class bcolors:
    PURPLE = '\033[95m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
