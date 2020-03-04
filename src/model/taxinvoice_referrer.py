import os
import hashlib

from bs4 import BeautifulSoup
import xlsxwriter

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_REFERRER, new_error,
                                  get_header_format, get_error_format)

HEADER_REFERRER = ['Commission Type', 'Client', 'Referrer Name', 'Amount Paid', 'GST Paid', 'Total Amount Paid']


class ReferrerTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.filetext = self.get_file_text()
        self.pair = None
        self.datarows = {}
        self.summary_errors = []
        self._key = self.__generate_key()
        self.parse()

    def get_file_text(self):
        file = open(self.full_path, 'r')
        return file.read()

    # region Parsers
    def parse(self):
        soup = BeautifulSoup(self.filetext, 'html.parser')

        self._from = self.parse_from(soup)
        self.from_abn = self.parse_from_abn(soup)
        self.to = self.parse_to(soup)
        self.to_abn = self.parse_to_abn(soup)
        self.bsb = self.parse_bsb(soup)
        self.account = self.parse_account(soup)
        self.final_total = self.parse_final_total(soup)
        self.datarows = self.parse_rows(soup)

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
        row_number = 0
        rows = {}
        for tr in table_rows:
            row_number += 1
            tds = tr.find_all('td')
            try:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                         tds[3].text, tds[4].text, tds[5].text, row_number)
                rows[row.key_full] = row
            except IndexError:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, '',
                                         tds[2].text, tds[3].text, tds[4].text, row_number)
                rows[row.key_full] = row
        return rows

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
    # endregion

    def __generate_key(self):
        sha = hashlib.sha256()

        filename_parts = self.filename.split('_')
        filename_parts = filename_parts[:-5]  # Remove process ID and date stamp

        for index, part in enumerate(filename_parts):
            if part == "Referrer":
                del filename_parts[index - 1]  # Remove year-month stamp

        filename_forkey = ''.join(filename_parts)
        sha.update(filename_forkey.encode(ENCODING))
        return sha.hexdigest()

    def process_comparison(self, margin=0.000001):
        if self.pair is None:
            return None
        assert type(self.pair) == type(self), "self.pair is not of the correct type"

        workbook = self.create_workbook()
        fmt_table_header = get_header_format(workbook)
        fmt_error = get_error_format(workbook)

        worksheet = workbook.add_worksheet()
        row = 0
        col_a = 0
        col_b = 8

        for index, item in enumerate(HEADER_REFERRER):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        # Code below is just to find the errors and write them into the spreadsheets
        for key in self.datarows.keys():
            self_row = self.datarows[key]
            pair_row = self.pair.datarows.get(key, None)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += ReferrerInvoiceRow.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right')

            self.summary_errors += ReferrerInvoiceRow.write_row(
                worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        alone_keys_infynity = set(self.pair.datarows.keys() - set(self.datarows.keys()))
        for key in alone_keys_infynity:
            self.summary_errors += ReferrerInvoiceRow.write_row(
                worksheet, self, self.pair.datarows[key], row, fmt_error, 'right')
            row += 1

        workbook.close()
        return self.summary_errors

    def create_workbook(self):
        suffix = '' if self.filename.endswith('.xlsx') else '.xlsx'
        return xlsxwriter.Workbook(OUTPUT_DIR_REFERRER + 'DETAILED_' + self.filename + suffix)


class ReferrerInvoiceRow(InvoiceRow):

    def __init__(self, commission_type, client, referrer, amount_paid, gst_paid, total, row_number):
        InvoiceRow.__init__(self)
        self._pair = None
        self._margin = 0

        self.commission_type = commission_type
        self.client = client
        self.referrer = referrer
        self.amount_paid = amount_paid
        self.gst_paid = gst_paid
        self.total = total

        self.row_number = row_number

        self._key = self.__generate_key()
        self._key_full = self.__generate_key_full()

    # region Properties
    @property
    def key(self):
        return self._key

    @property
    def key_full(self):
        return self._key_full

    @property
    def pair(self):
        return self._pair

    @pair.setter
    def pair(self, pair):
        self._pair = pair

    @property
    def margin(self):
        return self._margin

    @margin.setter
    def margin(self, margin):
        self._margin = margin

    @property
    def equal_commission_type(self):
        if self.pair is None:
            return False
        return self.commission_type == self.pair.commission_type

    @property
    def equal_client(self):
        if self.pair is None:
            return False
        return self.client == self.pair.client

    @property
    def equal_referrer(self):
        if self.pair is None:
            return False
        return self.referrer == self.pair.referrer

    @property
    def equal_amount_paid(self):
        if self.pair is None:
            return False
        return self.amount_paid == self.pair.amount_paid

    @property
    def equal_gst_paid(self):
        if self.pair is None:
            return False
        return self.gst_paid == self.pair.gst_paid

    @property
    def equal_total(self):
        if self.pair is None:
            return False
        return self.total == self.pair.total
    # endregion Properties

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

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left'):
        col = 0
        if side == 'right':
            col = 8

        worksheet.write(row, col, element.commission_type)
        worksheet.write(row, col + 1, element.client)
        worksheet.write(row, col + 2, element.referrer)

        format_ = fmt_error if not element.equal_amount_paid else None
        worksheet.write(row, col + 3, element.amount_paid, format_)

        format_ = fmt_error if not element.equal_gst_paid else None
        worksheet.write(row, col + 4, element.gst_paid, format_)

        format_ = fmt_error if not element.equal_total else None
        worksheet.write(row, col + 5, element.total, format_)

        errors = []
        line = element.row_number
        if element.pair is not None:

            if not element.equal_amount_paid:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Amount Paid does not match', line, element.amount_paid, element.pair.amount_paid))

            if not element.equal_gst_paid:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'GST Paid does not match', line, element.equal_gst_paid, element.pair.equal_gst_paid))

            if not element.equal_total:
                errors.append(new_error(
                    invoice.filename, invoice.pair.filename, 'Total does not match', line, element.total, element.pair.total))

        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line))

        return errors


def read_files_referrer(dir_: str, files: list) -> dict:
    keys = {}
    for file in files:
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = ReferrerTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
    return keys
