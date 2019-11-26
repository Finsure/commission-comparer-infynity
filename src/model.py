import copy
from bs4 import BeautifulSoup


class TaxInvoice:

    def __init__(self, directory, filename):
        self.directory = directory
        self.filename = filename
        self.filetext = self.get_file_text()
        self.parse()
        self._key = None

    def get_file_text(self):
        if self.directory[-1] != '/':
            self.directory += '/'
        file = open(self.directory + self.filename, 'r')
        return file.read()

    def parse(self):
        soup = BeautifulSoup(self.filetext, 'html.parser')

        self._from = self.parse_from(soup)
        self.from_abn = self.parse_from_abn(soup)
        self.to = self.parse_to(soup)
        self.to_abn = self.parse_to_abn(soup)
        self.bsb = self.parse_bsb(soup)
        self.account = self.parse_account(soup)
        self.final_total = self.parse_final_total(soup)
        self.rows = self.parse_rows(soup)

    def _get_parts_info(self, soup: BeautifulSoup):
        body = soup.find('body')
        extracted_info = body.find('p').text
        info = ' '.join(extracted_info.split())
        parts_info = info.split(':')
        return parts_info

    def _get_parts_account(self, soup: BeautifulSoup):
        body = soup.find('body')
        extracted_account = body.find('p').find_next('p').text
        account = ' '.join(extracted_account.split())
        parts_account = account.split(':')
        return parts_account

    def parse_from(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        _from = parts_info[1][:-4]
        _from = _from.strip()
        return _from

    def parse_from_abn(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        abn = parts_info[2][:-3]
        abn = abn.strip()
        return abn

    def parse_to(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        to = parts_info[3][:-4]
        to = to.strip()
        return to

    def parse_to_abn(self, soup: BeautifulSoup):
        parts_info = self._get_parts_info(soup)
        abn = parts_info[4][:-5]
        abn = abn.strip()
        return abn

    def parse_bsb(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        bsb = parts_account[1].split(' - ')[0].strip()
        return bsb

    def parse_account(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        account = parts_account[2].split('/')[0].strip()
        return account

    def parse_final_total(self, soup: BeautifulSoup):
        parts_account = self._get_parts_account(soup)
        final_total = parts_account[3].strip()
        return final_total

    def parse_rows(self, soup: BeautifulSoup):
        pass

    def key(self):
        if self._key is None:
            self.generate_key()
        return self._key

    def generate_key(self):
        self._key = ''
        pass

    def serialize(self):
        text = self.filetext
        self.filetext = None
        serialized_obj = copy.copy(self.__dict__)
        self.filetext = text
        return serialized_obj
