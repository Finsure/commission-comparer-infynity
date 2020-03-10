import os
import hashlib

from bs4 import BeautifulSoup
import xlsxwriter

from src.model.taxinvoice import (TaxInvoice, InvoiceRow, ENCODING, OUTPUT_DIR_REFERRER, new_error,
                                  get_header_format, get_error_format)

from src import utils as u

HEADER_REFERRER = ['Commission Type', 'Client', 'Referrer Name', 'Amount Paid', 'GST Paid', 'Total Amount Paid']


class ReferrerTaxInvoice(TaxInvoice):

    def __init__(self, directory, filename):
        TaxInvoice.__init__(self, directory, filename)
        self.filetext = self.get_file_text()
        self.pair = None
        self.datarows = {}
        self.datarows_count = {}
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
        self.parse_rows(soup)

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
        for tr in table_rows:
            row_number += 1
            tds = tr.find_all('td')
            try:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, tds[2].text,
                                         tds[3].text, tds[4].text, tds[5].text, row_number)
                self.__add_datarow(row)
            except IndexError:
                row = ReferrerInvoiceRow(tds[0].text, tds[1].text, '',
                                         tds[2].text, tds[3].text, tds[4].text, row_number)
                self.__add_datarow(row)

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

        format_ = fmt_error if not self.equal_from else None
        worksheet.write(row, col_a, 'From')
        worksheet.write(row, col_a + 1, self._from, format_)
        row += 1
        format_ = fmt_error if not self.equal_from_abn else None
        worksheet.write(row, col_a, 'From ABN')
        worksheet.write(row, col_a + 1, self.from_abn, format_)
        row += 1
        format_ = fmt_error if not self.equal_to else None
        worksheet.write(row, col_a, 'To')
        worksheet.write(row, col_a + 1, self.to, format_)
        row += 1
        format_ = fmt_error if not self.equal_to_abn else None
        worksheet.write(row, col_a, 'To ABN')
        worksheet.write(row, col_a + 1, self.to_abn, format_)
        row += 1
        format_ = fmt_error if not self.equal_bsb else None
        worksheet.write(row, col_a, 'BSB')
        worksheet.write(row, col_a + 1, self.bsb, format_)
        row += 1
        format_ = fmt_error if not self.equal_account else None
        worksheet.write(row, col_a, 'Account')
        worksheet.write(row, col_a + 1, self.account, format_)
        row += 1
        format_ = fmt_error if not self.equal_final_total else None
        worksheet.write(row, col_a, 'Total')
        worksheet.write(row, col_a + 1, self.final_total, format_)

        if self.pair is not None:
            row = 0
            format_ = fmt_error if not self.equal_from else None
            worksheet.write(row, col_b, 'From')
            worksheet.write(row, col_b + 1, self._from, format_)
            row += 1
            format_ = fmt_error if not self.equal_from_abn else None
            worksheet.write(row, col_b, 'From ABN')
            worksheet.write(row, col_b + 1, self.from_abn, format_)
            row += 1
            format_ = fmt_error if not self.equal_to else None
            worksheet.write(row, col_b, 'To')
            worksheet.write(row, col_b + 1, self.to, format_)
            row += 1
            format_ = fmt_error if not self.equal_to_abn else None
            worksheet.write(row, col_b, 'To ABN')
            worksheet.write(row, col_b + 1, self.to_abn, format_)
            row += 1
            format_ = fmt_error if not self.equal_bsb else None
            worksheet.write(row, col_b, 'BSB')
            worksheet.write(row, col_b + 1, self.bsb, format_)
            row += 1
            format_ = fmt_error if not self.equal_account else None
            worksheet.write(row, col_b, 'Account')
            worksheet.write(row, col_b + 1, self.account, format_)
            row += 1
            format_ = fmt_error if not self.equal_final_total else None
            worksheet.write(row, col_b, 'Total')
            worksheet.write(row, col_b + 1, self.final_total, format_)

        row += 2

        for index, item in enumerate(HEADER_REFERRER):
            worksheet.write(row, col_a + index, item, fmt_table_header)
            worksheet.write(row, col_b + index, item, fmt_table_header)
        row += 1

        keys_unmatched = set(self.pair.datarows.keys() - set(self.datarows.keys()))

        # Code below is just to find the errors and write them into the spreadsheets
        for key_full in self.datarows.keys():
            self_row = self.datarows[key_full]
            pair_row = self.find_pair_row(self_row)

            self_row.margin = margin
            self_row.pair = pair_row

            if pair_row is not None:
                # delete from pair list so it doesn't get matched again
                del self.pair.datarows[pair_row.key_full]
                # Remove the key from the keys_unmatched if it is there
                if pair_row.key_full in keys_unmatched:
                    keys_unmatched.remove(pair_row.key_full)

                pair_row.margin = margin
                pair_row.pair = self_row
                self.summary_errors += ReferrerInvoiceRow.write_row(
                    worksheet, self, pair_row, row, fmt_error, 'right', write_errors=False)

            self.summary_errors += ReferrerInvoiceRow.write_row(worksheet, self, self_row, row, fmt_error)
            row += 1

        # Write unmatched records
        for key in keys_unmatched:
            self.summary_errors += ReferrerInvoiceRow.write_row(
                worksheet, self, self.pair.datarows[key], row, fmt_error, 'right')
            row += 1

        if len(self.summary_errors) > 0:
            workbook.close()
        else:
            del workbook
        return self.summary_errors

    def find_pair_row(self, row):
        # Match by full_key
        pair_row = self.pair.datarows.get(row.key_full, None)
        if pair_row is not None:
            return pair_row

        # We want to match by similarity before matching by the key
        # Match by similarity
        for key, item in self.pair.datarows.items():
            if row.equals(item):
                return item

        # Match by key
        for key, item in self.pair.datarows.items():
            if row.key == item.key:
                return item

        # Return None if nothing found
        return None

    def create_workbook(self):
        suffix = '' if self.filename.endswith('.xlsx') else '.xlsx'
        return xlsxwriter.Workbook(OUTPUT_DIR_REFERRER + 'DETAILED_' + self.filename + suffix)

    def __add_datarow(self, row):
        if row.key in self.datarows.keys():  # If the row already exists
            self.datarows_count[row.key] += 1  # Increment row count for that key
            row.key = row._generate_key(self.datarows_count[row.key])  # Generate new key for the record
            self.datarows[row.key] = row  # Add row to the list
        else:
            self.datarows_count[row.key] = 0  # Start counter
            self.datarows[row.key] = row  # Add row to the list

    # region Properties
    @property
    def equal_from(self):
        if self.pair is None:
            return False
        return self._from == self.pair._from

    @property
    def equal_from_abn(self):
        if self.pair is None:
            return False
        return self.from_abn == self.pair.from_abn

    @property
    def equal_to(self):
        if self.pair is None:
            return False
        return self.to == self.pair.to

    @property
    def equal_to_abn(self):
        if self.pair is None:
            return False
        return self.to_abn == self.pair.to_abn

    @property
    def equal_bsb(self):
        if self.pair is None:
            return False
        return self.bsb == self.pair.bsb

    @property
    def equal_account(self):
        if self.pair is None:
            return False
        return self.account == self.pair.account

    @property
    def equal_final_total(self):
        if self.pair is None:
            return False
        return self.final_total == self.pair.final_total
    # endregion


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

        self._key = self._generate_key()
        self._key_full = self.__generate_key_full()

    # region Properties
    @property
    def key(self):
        return self._key

    @key.setter
    def key(self, k):
        self._key = k

    @property
    def key_full(self):
        return self._key_full

    @key_full.setter
    def key_full(self, k):
        self._key_full = k

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
        return u.sanitize(self.commission_type) == u.sanitize(self.pair.commission_type)

    @property
    def equal_client(self):
        if self.pair is None:
            return False
        return u.sanitize(self.client) == u.sanitize(self.pair.client)

    @property
    def equal_referrer(self):
        if self.pair is None:
            return False
        return u.sanitize(self.referrer) == u.sanitize(self.pair.referrer)

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

    def _generate_key(self, salt=''):
        sha = hashlib.sha256()
        sha.update(u.sanitize(self.commission_type).encode(ENCODING))
        sha.update(u.sanitize(self.client).encode(ENCODING))
        sha.update(u.sanitize(self.referrer).encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def __generate_key_full(self, salt=''):
        sha = hashlib.sha256()
        sha.update(self.commission_type.encode(ENCODING))
        sha.update(self.client.encode(ENCODING))
        sha.update(self.referrer.encode(ENCODING))
        sha.update(self.amount_paid.encode(ENCODING))
        sha.update(self.gst_paid.encode(ENCODING))
        sha.update(self.total.encode(ENCODING))
        sha.update(str(salt).encode(ENCODING))
        return sha.hexdigest()

    def equals(self, obj):
        if type(obj) != ReferrerInvoiceRow:
            return False

        return (
            u.sanitize(self.commission_type) == u.sanitize(obj.commission_type)
            and u.sanitize(self.client) == u.sanitize(obj.client)
            and u.sanitize(self.referrer) == u.sanitize(obj.referrer)
            and self.compare_numbers(self.amount_paid, obj.amount_paid, self.margin)
            and self.compare_numbers(self.gst_paid, obj.gst_paid, self.margin)
            and self.compare_numbers(self.total, obj.total, self.margin)
        )

    @staticmethod
    def write_row(worksheet, invoice, element, row, fmt_error, side='left', write_errors=True):
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
            if write_errors:
                if not element.equal_amount_paid:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Amount Paid does not match', line, element.amount_paid, element.pair.amount_paid))

                if not element.equal_gst_paid:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'GST Paid does not match', line, element.gst_paid, element.pair.gst_paid))

                if not element.equal_total:
                    errors.append(new_error(
                        invoice.filename, invoice.pair.filename, 'Total does not match', line, element.total, element.pair.total))

        else:
            errors.append(new_error(invoice.filename, invoice.pair.filename, 'No corresponding row in commission file', line))

        return errors


def read_files_referrer(dir_: str, files: list) -> dict:
    keys = {}
    counter = 1
    for file in files:
        print(f'Parsing {counter} of {len(files)} files from {dir_}', end='\r')
        if os.path.isdir(dir_ + file):
            continue
        try:
            ti = ReferrerTaxInvoice(dir_, file)
            keys[ti.key] = ti
        except IndexError:
            # handle exception when there is a column missing in the file.
            pass
        counter += 1
    print()
    return keys
