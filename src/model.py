import copy
import hashlib
from bs4 import BeautifulSoup


ENCODING = 'utf-8'


class TaxInvoice:

    def __init__(self, directory, filename):
        self.directory = directory
        self.filename = filename
        self.filetext = self.get_file_text()
        self.parse()

        self._key = None

    def get_full_path(self):
        self.__fix_path()
        return self.directory + self.filename

    def get_file_text(self):
        file = open(self.get_full_path, 'r')
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
        header = soup.find('tr')  # Find header
        header.extract()  # Remove header
        table_rows = soup.find_all('tr')
        rows = []
        for tr in table_rows:
            tds = tr.find_all('td')
            try:
                row = InvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                 tds[3].text, tds[4].text, tds[5].text)
                rows.append(row)
            except IndexError:
                row = InvoiceRow(tds[0].text, tds[1].text, '',
                                 tds[2].text, tds[3].text, tds[4].text)
                rows.append(row)
        return rows

    def key(self):
        if self._key is None:
            self._key = self.__generate_key()
        return self._key

    def serialize(self):
        text = self.filetext
        self.filetext = None
        serialized_obj = copy.copy(self.__dict__)
        self.filetext = text
        return serialized_obj

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

    def __fix_path(self):
        if self.directory[-1] != '/':
            self.directory += '/'

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(self.filename.encode(ENCODING))
        return sha.hexdigest()


class InvoiceRow:

    def __init__(self, commission_type, client, referrer, amount_paid, gst_paid, total):
        self.commission_type = commission_type
        self.client = client
        self.referrer = referrer
        self.amount_paid = amount_paid
        self.gst_paid = gst_paid
        self.total = total

        self._key = None
        self._key_full = None

    def key(self):
        if self._key is None:
            self._key = self.__generate_key()
        return self._key

    def key_full(self):
        if self._key_full is None:
            self._key_full = self.__generate_key_full()
        return self._key_full

    def __generate_key(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.referrer.encode(ENCODING))
        return sha.hexdigest()

    def __generate_key_full(self):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.referrer.encode(ENCODING))
        sha.update(self.amount_paid.encode(ENCODING))
        sha.update(self.gst_paid.encode(ENCODING))
        sha.update(self.total.encode(ENCODING))
        return sha.hexdigest()

    def serialize(self):
        return self.__dict__
