import copy
import hashlib

from bs4 import BeautifulSoup

from src.utils import merge_lists

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
        file = open(self.get_full_path(), 'r')
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
        rows = {}
        for tr in table_rows:
            tds = tr.find_all('td')
            try:
                row = InvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                 tds[3].text, tds[4].text, tds[5].text)
                rows[row.key_full()] = row
            except IndexError:
                row = InvoiceRow(tds[0].text, tds[1].text, '',
                                 tds[2].text, tds[3].text, tds[4].text)
                rows[row.key_full()] = row
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

    def compare_to(self, invoice, margin=0.0000001):  # noqa F821
        result = result_invoice()
        result['filename'] = self.filename

        if invoice is None:
            result['invoice'] = self.get_full_path()
            return result

        #  If we reached here it means the file has a pair
        result['has_pair'] = True

        # ensure these have been parsed
        if len(self.rows) == 0:
            self.parse()
        if len(invoice.rows) == 0:
            invoice.parse()

        keys_all = merge_lists(self.rows.keys(), invoice.rows.keys())

        result_rows = {}

        for key in keys_all:
            row_local = self.rows.get(key, None)
            row_invoice = invoice.rows.get(key, None)

            # If we couldnt find the row by the InvoiceRow.full_key() it means they are different
            # so we try to locate them by the InvoiceRow.key()
            if row_local is None:
                for k in self.rows.keys():
                    if self.rows.get(k, None).key() == row_invoice.key():
                        row_local = self.rows[k]
                        keys_all.remove(row_invoice.key_full())
            elif row_invoice is None:
                for k in invoice.rows.keys():
                    if invoice.rows.get(k, None).key() == row_local.key():
                        row_invoice = invoice.rows[k]
                        keys_all.remove(row_local.key_full())

            if row_local is not None:
                result_rows[key] = row_local.compare_to(row_invoice)
            else:
                result_rows[key] = row_invoice.compare_to(row_local)

        result['results_rows'] = result_rows

        return result

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

    def serialize(self):
        return self.__dict__

    def compare_to(self, row, margin=0.0000001):  # noqa F821
        if row is None:
            return {
                'overall': False,
                'commission_type': False,
                'client': False,
                'referrer': False,
                'amount_paid': False,
                'gst_paid': False,
                'total': False
            }
        equal_commission_type = self.commission_type == row.commission_type
        equal_client = self.client == row.client
        equal_referrer = self.referrer == row.referrer
        equal_amount_paid = self.amount_paid == row.amount_paid
        equal_gst_paid = self.gst_paid == row.gst_paid
        equal_total = self.total == row.total

        # Recompare monetary values using the
        if not equal_amount_paid:
            equal_amount_paid = self.compare_numbers(self.amount_paid, row.amount_paid, margin)
        if not equal_gst_paid:
            equal_gst_paid = self.compare_numbers(self.gst_paid, row.gst_paid, margin)
        if not equal_total:
            equal_total = self.compare_numbers(self.total, row.total, margin)

        overall = (equal_commission_type and equal_client and equal_referrer
                   and equal_amount_paid and equal_gst_paid and equal_total)

        return {
            'overall': overall,
            'commission_type': equal_commission_type,
            'client': equal_client,
            'referrer': equal_referrer,
            'amount_paid': equal_amount_paid,
            'gst_paid': equal_gst_paid,
            'total': equal_total
        }

    def compare_numbers(self, n1, n2, margin):
        n1val = n1
        n2val = n2

        if n1 or n2 == '':
            return False

        if type(n1) == str:
            n1val = float(n1[-1:])  # remove $
        if type(n2) == str:
            n2val = float(n2[-1:])  # remove $

        return abs(n1val - n2val) <= margin

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


def result_invoice():
    return {
        'filename': '',
        'has_pair': False,
        'results_rows': '',
    }


def result_row():
    return {

    }
