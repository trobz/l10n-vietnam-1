# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from datetime import datetime
from dateutil.relativedelta import relativedelta
import logging
from openerp.addons.report_xlsx.report.report_xlsx \
    import ReportXlsx  # @UnresolvedImport
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF


_logger = logging.getLogger(__name__)


class CashFlowDirectReport(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]

        self.row_pos = 0
        self._define_formats(workbook)

        self.sheet = workbook.add_worksheet(name="Cash Flow (Direct)")
        # Hide grid of the file
        self.sheet.hide_gridlines(2)
        self.format_data_sheet()
        self._set_cells_size()
        self.write_title()
        self.write_report_info()
        self.write_report_headers()
        self.generate_main_content()
        self.write_signature_section()

    def write_signature_section(self):
        row = self.row + 2
        self.sheet.write(
            "A%s" % row,
            u"Người lập biểu",
            self.fm_bold_center
        )

        self.sheet.write(
            "A%s" % (row + 1),
            u"(Ký, họ tên)",
            self.fm_center
        )

        self.sheet.merge_range(
            "B%s:C%s" % (row, row),
            u"Kế toán trưởng",
            self.fm_bold_center
        )

        self.sheet.merge_range(
            "B%s:C%s" % (row + 1, row + 1),
            u"(Ký, họ tên)",
            self.fm_center
        )

        self.sheet.merge_range(
            "D%s:E%s" % (row, row),
            u"Giám đốc",
            self.fm_bold_center
        )

        self.sheet.merge_range(
            "D%s:E%s" % (row + 1, row + 1),
            u"(Ký, họ tên, đóng dấu)",
            self.fm_center
        )

    def format_data_sheet(self):
        self.sheet.set_portrait()
        self.sheet.set_paper(9)
        self.sheet.fit_to_pages(1, 0)

    def write_title(self):
        company_info = self.object.env.user.company_id.get_company_info()
        company_name = company_info.get('name', '')
        company_address = company_info.get('address', '')
        self.sheet.merge_range(
            "A1:B1",
            u"Đơn vị báo cáo: %s" % company_name,
            self.fm_bold
        )
        self.sheet.merge_range(
            "A2:B2",
            u"Địa chỉ: %s" % company_address,
            self.fm_bold
        )
        self.sheet.merge_range(
            "A3:B3",
            u"MST: ",
            self.fm_bold
        )
        self.sheet.merge_range(
            "C1:E1",
            u"Mẫu số B 03 - DN",
            self.fm_bold_center
        )
        self.sheet.merge_range(
            "C2:E2",
            u"(Ban hành theo Thông tư số 200/2014/TT-BTC, "
            u"Ngày 22/12/2014 của BTC))",
            self.fm_center
        )

    def write_report_info(self):
        self.sheet.merge_range(
            "A5:E5",
            u"BÁO CÁO LƯU CHUYỂN TIỀN TỆ",
            self.fm_report_title
        )
        self.sheet.merge_range(
            "A6:E6",
            u"(Theo phương pháp trực tiếp) (*)",
            self.fm_center
        )
        self.sheet.merge_range(
            "A7:E7",
            u"Từ %s đến %s" % (
                self.convert_date_format(self.object.date_from),
                self.convert_date_format(self.object.date_to)
            ),
            self.fm_italic_center
        )

    def convert_date_format(self, date_string):
        date = datetime.strptime(date_string, DF)
        return datetime.strftime(date, "%d/%m/%Y")

    def write_report_headers(self):
        headers = {
            0: u"Chỉ tiêu",
            1: u"Mã số",
            2: u"Thuyết minh",
            3: u"Năm nay",
            4: u"Năm trước",
        }
        for index, header in headers.iteritems():
            self.sheet.write(
                8, index, header, self.fm_tb_bold_center
            )
            self.sheet.write(
                9, index, index + 1, self.fm_tb_bold_center
            )

    def get_data(self):
        state_condition = self.object.target_move == "posted" and \
            "AND MOVE.state = 'posted'" or ''

        company_condition = self.object.company_id and \
            ("AND LINE.company_id = %(c_id)s "
             "AND COUNTER_LINE.company_id = %(c_id)s") % {
                'c_id':  self.object.company_id.id
            } or ''

        self.params.update(
            {
                "state_condition": state_condition,
                "company_condition": company_condition
            }
        )

        sql = '''
        SELECT
            LINE.move_id,
            CASE
                WHEN
                    LINE.date < '%(date_from_last_period)s'
                THEN TRUE ELSE FALSE
            END AS is_before_last_period,
            CASE
                WHEN
                    LINE.date BETWEEN '%(date_from_last_period)s' AND
                    '%(date_to_last_period)s'
                THEN TRUE ELSE FALSE
            END AS is_last_period,
            CASE
                WHEN
                    LINE.date < '%(date_from_this_period)s'
                THEN TRUE ELSE FALSE
            END AS is_before_this_period,
            CASE
                WHEN
                    LINE.date BETWEEN '%(date_from_this_period)s' AND
                    '%(date_to_this_period)s'
                THEN TRUE ELSE FALSE
            END AS is_this_period,

            LINE_ACCOUNT.code AS code,
            LINE.debit AS debit,
            LINE.credit AS credit,
            COUNTER_LINE_ACCOUNT.code AS counter_code,
            LINE.debit AS counter_credit,
            LINE.credit AS counter_debit

            /*
                Take counter_credit = LINE.debit because every single line of
                    selected table is a pair of mutual journal items.
                    So, debit must be equal to credit. If it is not,
                    the generated/created journal entry is incorrect.

                Don't take the counter_credit = COUNTER_LINE.credit
                    because LINE is the journal item WITH counterpart
                    and COUNTER_LINE is the one WITHOUT counterpart
                    There will be POSSIBLE for multiple LINEs to have
                    the same COUNTER_LINE, then COUNTER_LINE.debit/credit
                    will be the sum of ALL LINE.credit/debit
            */

        FROM
            account_move_line LINE
            LEFT JOIN account_account LINE_ACCOUNT
                ON LINE_ACCOUNT.id = LINE.account_id
            LEFT JOIN account_move_line COUNTER_LINE
                ON COUNTER_LINE.id = LINE.counter_move_id
            LEFT JOIN account_account COUNTER_LINE_ACCOUNT
                ON COUNTER_LINE_ACCOUNT.id = COUNTER_LINE.account_id
            LEFT JOIN account_move MOVE
                ON MOVE.id = LINE.move_id

        WHERE
            LINE.counter_move_id IS NOT NULL AND
            LINE.date <= '%(date_to_this_period)s'
            %(state_condition)s
            %(company_condition)s

        ORDER BY
            move_id
        ''' % self.params

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()

        return data

    def get_date_params(self):
        date_from = datetime.strptime(self.object.date_from, DF)
        date_to = datetime.strptime(self.object.date_to, DF)
        month_interval = relativedelta(date_to, date_from).months

        date_from_last_year = date_from + relativedelta(
            months=-(1 + month_interval), day=1)

        date_to_last_year = date_to + relativedelta(
            months=-(1 + month_interval), day=31)

        self.params = {
            'date_from_this_period': date_from.date(),
            'date_to_this_period': date_to.date(),

            'date_from_last_period': date_from_last_year.date(),
            'date_to_last_period': date_to_last_year.date(),
        }

    def generate_main_content(self):
        self.configs = self.object.env['cash.flow.direct.config'].search(
            [], order='sequence'
        )
        self.get_date_params()
        self.data = self.get_data()
        report_data = self.prepare_report_data()
        if len(report_data):
            self.row = 11

            # A dictionary with key: value coresponding to
            #        {config_id: xlsx_file_row}
            self.line_data = {}
            self.write_raw_data_lines(report_data)

    def set_row_height(self, row, description):
        number_description_line = len(description.split("\n"))
        self.sheet.set_row(row, 18 * number_description_line)
        return

    def write_total_lines(self, report_data):
        for row, line in report_data.iteritems():
            self.line_data[line['id']] = row

            is_bold = line['bold']
            description = line['description'].strip()
            child_ids = line['child_ids']
            this_period_formula, last_period_formula = \
                self.get_sum_formula(child_ids)

            self.sheet.write(
                'A%s' % row,
                line['name'],
                is_bold and self.fm_tb_bold or self.fm_tb_default
            )

            self.sheet.write(
                'B%s' % row,
                line['code'] or '',
                is_bold and self.fm_tb_bold_center or self.fm_tb_center
            )

            self.sheet.write(
                'C%s' % row,
                description,
                is_bold and self.fm_tb_bold or self.fm_tb_default
            )

            self.sheet.write_formula(
                'D%s' % row,
                this_period_formula,
                is_bold and self.fm_tb_number_bold or self.fm_tb_number
            )

            self.sheet.write_formula(
                'E%s' % row,
                last_period_formula,
                is_bold and self.fm_tb_number_bold or self.fm_tb_number
            )

            self.set_row_height(row - 1, description)

    def get_correct_amount(self, amount, line_type, has_parenthesis):
        amount = amount is not None and amount or ''
        if not (isinstance, (int, float)):
            return amount

        amount = amount > 0 and has_parenthesis and (amount * -1) or amount
        if not amount:
            if line_type == 'detail_line':
                amount = 0.0
            else:
                amount = ''

        return amount

    def write_raw_data_lines(self, report_data):
        # We will write total_line after writing all other lines
        #    in order to make sure their id are listed in self.line_data
        total_lines = {}

        for line in report_data:
            self.line_data[line['id']] = self.row
            has_parenthesis = line['parenthesis']
            description = line['description'].strip()
            is_bold = line['bold']
            line_type = line['type']

            if line_type == 'total_line':
                # We reserved this row for the total_line
                total_lines[self.row] = line
                self.row += 1
                continue

            amount_this_period = line.get('amount_this_period', None)
            amount_this_period = self.get_correct_amount(
                amount_this_period, line_type, has_parenthesis)

            amount_last_period = line.get('amount_last_period', None)
            amount_last_period = self.get_correct_amount(
                amount_last_period, line_type, has_parenthesis)

            self.sheet.write(
                'A%s' % self.row,
                line['name'],
                is_bold and self.fm_tb_bold or self.fm_tb_default
            )

            self.sheet.write(
                'B%s' % self.row,
                line['code'] or '',
                is_bold and self.fm_tb_bold_center or self.fm_tb_center
            )

            self.sheet.write(
                'C%s' % self.row,

                description,
                is_bold and self.fm_tb_bold or self.fm_tb_default
            )

            self.sheet.write(
                'D%s' % self.row,
                amount_this_period,
                is_bold and self.fm_tb_number_bold or self.fm_tb_number
            )

            self.sheet.write(
                'E%s' % self.row,
                amount_last_period,
                is_bold and self.fm_tb_number_bold or self.fm_tb_number
            )

            self.set_row_height(self.row - 1, description)
            self.row += 1

        self.write_total_lines(total_lines)

    def get_sum_formula(self, child_ids):
        if not child_ids or not isinstance(child_ids, list):
            return ("=SUM(0)", "=SUM(0)")

        try:
            data_row = [self.line_data[row] for row in child_ids]
            this_period_cells = ['D%s' % row for row in data_row]
            this_period_formula = len(this_period_cells) and \
                "=SUM(%s)" % ",".join(this_period_cells) or "=SUM(0)"

            last_period_cells = ['E%s' % row for row in data_row]
            last_period_formula = len(last_period_cells) and \
                "=SUM(%s)" % ",".join(last_period_cells) or "=SUM(0)"

            return (this_period_formula, last_period_formula)

        except:
            _logger.error(
                "Error occurred during printing Cash Flow (Direct) Report."
            )
            return ("=SUM(0)", "=SUM(0)")

    def prepare_report_data(self):
        report_data = []

        self.is_before = False
        for config in self.configs:

            self.is_before = config.data_period == 'before'
            line = {
                'name': config.name,
                'bold': config.is_bold,
                'parenthesis': config.has_parenthesis,
                'code': config.code or '',
                'description': config.description or '',
                'type': config.type,
                'id': config.id,
                'child_ids': config.child_config_ids.ids
            }

            if config.type == "detail_line":

                line.update(
                    self.get_detail_line_data(config)
                )

            report_data.append(line)

        return report_data

    def get_detail_line_data(self, config):
        data_method = config.data_method
        detail_line_data = {}

        # Take all account_move_line, just care about the credit accounts
        if data_method == 'from_credit':
            cr_accounts = self.get_account_list(config.credit_accounts)
            detail_line_data = self.get_credit_debit_only(
                cr_accounts, credit=True)

        # Take all account_move_line, just care about the debit accounts
        elif data_method == 'from_debit':
            dr_accounts = self.get_account_list(config.debit_accounts)
            detail_line_data = self.get_credit_debit_only(
                dr_accounts)

        # Take account_move_line, which matches both debit and credit acc.
        elif data_method == 'from_both':
            detail_line_data = self.get_credit_and_debit(config)

        # Take account_move_line, that matches either debit or credit acc.
        elif data_method == 'from_either':
            detail_line_data = self.get_credit_or_debit(config)

        return detail_line_data

    def get_account_list(self, account_string):
        if not account_string or not (
            isinstance(account_string, bool) or account_string.strip()
        ):
            return []

        return [
            a.strip() for a in account_string.split(",")]

    def get_credit_and_debit_amount(self, cr_accounts, dr_accounts):

        cr_accounts = tuple(cr_accounts)
        dr_accounts = tuple(dr_accounts)

        this_period, last_period = self.is_before and\
            ('is_before_this_period', 'is_before_last_period') or \
            ('is_this_period', 'is_last_period')

        line_data = {
            'amount_this_period': 0,
            'amount_last_period': 0
        }
        for data in self.data:
            # Don't care about credit/debit or counter_debit/counter_credit
            # Because they are equal, if not, journal entry is incorrect.
            amount = data['credit'] or data['debit']

            if (
                data['credit'] and
                data['code'].startswith(cr_accounts) and
                data['counter_code'].startswith(dr_accounts)
            ) or (
                data['debit'] and
                data['code'].startswith(dr_accounts) and
                data['counter_code'].startswith(cr_accounts)
            ):

                if data[last_period]:
                    line_data["amount_last_period"] += amount
                elif data[this_period]:
                    line_data["amount_this_period"] += amount

        return line_data

    def get_credit_and_debit(self, config):
        compute_method = config.compute_method
        cr_accounts = self.get_account_list(
            config.credit_accounts)
        dr_accounts = self.get_account_list(
            config.debit_accounts)

        if compute_method in ['sum_debit', 'sum_credit']:
            # Get all Dr ammount from account.move.line of given
            #     debit accounts AND credit accounts
            return self.get_credit_and_debit_amount(
                cr_accounts, dr_accounts)

        else:
            # Involving Accounts to compute Total Credit
            adv_cr_accounts = self.get_account_list(
                config.advanced_credit_accounts)
            adv_dr_accounts = self.get_account_list(
                config.advanced_debit_accounts)

            total_credit = self.get_credit_and_debit_amount(
                adv_cr_accounts, adv_dr_accounts
            )
            total_debit = self.get_credit_and_debit_amount(
                cr_accounts, dr_accounts
            )

            if compute_method == 'debit_credit':
                return {
                    key: total_debit[key] - total_credit[key]
                    for key in total_credit.keys()
                }

            return {
                key: total_credit[key] - total_debit[key]
                for key in total_credit.keys()
            }

    def get_credit_or_debit(self, config):
        compute_method = config.compute_method

        cr_accounts = self.get_account_list(config.credit_accounts)
        dr_accounts = self.get_account_list(config.debit_accounts)

        debit_data = self.get_credit_debit_only(dr_accounts)
        credit_data = self.get_credit_debit_only(cr_accounts, credit=True)

        if compute_method == 'sum_debit':
            # Get all Dr ammount from account.move.line of given accounts
            # Actually, this is the same with data_method == 'from_debit'
            return debit_data

        elif compute_method == 'sum_credit':
            # Get all Cr ammount from account.move.line of given accounts
            # Actually, this is the same with data_method == 'from_credit'
            return credit_data

        elif compute_method == 'credit_debit':
            return {
                key: credit_data[key] - debit_data[key]
                for key in credit_data.keys()
            }

        elif compute_method == 'debit_credit':
            return {
                key: debit_data[key] - credit_data[key]
                for key in credit_data.keys()
            }

    def get_credit_debit_only(self, accounts, credit=False):
        '''
            This function will get all Cr/Dr ammount from account.move.line
                of given accounts
        '''
        this_period, last_period = self.is_before and\
            ('is_before_this_period', 'is_before_last_period') or \
            ('is_this_period', 'is_last_period')

        accounts = tuple(accounts)
        amount_col, counter_amount_col = credit and (
            'credit', 'counter_credit') or ('debit', 'counter_debit')

        line_data = {
            'amount_this_period': 0,
            'amount_last_period': 0
        }
        for data in self.data:
            amount = data[amount_col] or data[counter_amount_col]

            if (
                data[amount_col] and
                data['code'].startswith(accounts)
            ) or (
                data[counter_amount_col] and
                data['counter_code'].startswith(accounts)
            ):

                if data[last_period]:
                    line_data["amount_last_period"] += amount
                elif data[this_period]:
                    line_data["amount_this_period"] += amount

        return line_data

    def _set_cells_size(self):
        # Set default height
        self.sheet.set_default_row(18)
        # Set default format
        self.sheet.set_column('A:A', 70, self.fm_default)
        self.sheet.set_column('B:B', 10, self.fm_default)
        self.sheet.set_column('C:C', 30, self.fm_default)
        self.sheet.set_column('D:E', 20, self.fm_default)

    def create_format(self, format_dict):
        return self.workbook.add_format(format_dict)

    def _define_formats(self, workbook):
        """
        Add cell formats to current workbook.
        Those formats can be used on all cell.
        """
        self.workbook = workbook
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        fm_default = {
            'font_name': 'Arial',
            'font_size': 10,
            'valign': 'vcenter',
            'num_format': '#,##0',
        }
        self.fm_default = self.create_format(fm_default)
        # ---------------------------------------------------------------------
        fm_center = fm_default.copy()
        fm_center.update({
            'align': 'center',
        })
        self.fm_center = self.create_format(fm_center)
        # ---------------------------------------------------------------------
        fm_bold = fm_default.copy()
        fm_bold.update({
            'bold': True,
            'align': 'left',
        })
        self.fm_bold = self.create_format(fm_bold)
        # ---------------------------------------------------------------------
        fm_italic = fm_default.copy()
        fm_italic.update({
            'italic': True,
            'align': 'left',
        })
        self.fm_italic = self.create_format(fm_italic)
        # ---------------------------------------------------------------------
        fm_italic_center = fm_italic.copy()
        fm_italic_center.update({
            'italic': True,
            'align': 'center',
        })
        self.fm_italic_center = self.create_format(fm_italic_center)
        # ---------------------------------------------------------------------
        fm_bold_center = fm_bold.copy()
        fm_bold_center.update({
            'bold': True,
            'align': 'center',
        })
        self.fm_bold_center = self.create_format(fm_bold_center)
        # ---------------------------------------------------------------------
        fm_bold_center_italic = fm_bold_center.copy()
        fm_bold_center_italic.update({
            'bold': True,
            'align': 'center',
            'italic': True
        })
        self.fm_bold_center_italic = self.create_format(fm_bold_center_italic)
        # ---------------------------------------------------------------------
        fm_report_title = fm_default.copy()
        fm_report_title.update({
            'bold': True,
            'align': 'center',
            'font_size': 18,
        })
        self.fm_report_title = self.create_format(fm_report_title)

        # ---------------------------------------------------------------------
        # Table
        # ---------------------------------------------------------------------
        fm_tb_default = fm_default.copy()
        fm_tb_default.update({
            'border': True,
            'align': 'left',
            'text_wrap': True
        })
        self.fm_tb_default = self.create_format(fm_tb_default)
        # ---------------------------------------------------------------------
        fm_tb_number = fm_tb_default.copy()
        fm_tb_number.update({
            'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        })
        self.fm_tb_number = self.create_format(fm_tb_number)
        # ---------------------------------------------------------------------
        fm_tb_number_bold = fm_tb_number.copy()
        fm_tb_number_bold.update({
            'bold': True,
        })
        self.fm_tb_number_bold = self.create_format(fm_tb_number_bold)
        # ---------------------------------------------------------------------
        fm_tb_bold = fm_tb_default.copy()
        fm_tb_bold.update({
            'border': True,
            'bold': True,
            'align': 'left',
        })
        self.fm_tb_bold = self.create_format(fm_tb_bold)
        # ---------------------------------------------------------------------
        fm_tb_center = fm_tb_default.copy()
        fm_tb_center.update({
            'border': True,
            'align': 'center',
        })
        self.fm_tb_center = self.create_format(fm_tb_center)
        # ---------------------------------------------------------------------
        fm_tb_bold_center = fm_tb_bold.copy()
        fm_tb_bold_center.update({
            'border': True,
            'bold': True,
            'align': 'center',
        })
        self.fm_tb_bold_center = self.create_format(fm_tb_bold_center)


CashFlowDirectReport(
    'report.report_cash_flow_direct_xlsx', 'account.cash.flow.report.wizard')
