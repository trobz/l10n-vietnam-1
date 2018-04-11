# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from datetime import datetime
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF


class CashBankBookReportXlsx(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]
        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self.setup_config()

        # generate header
        self.generate_company_info()
        self.generate_report_template()

        # generate main content
        self.generate_report_title()
        self.generate_content_header()
        self.generate_main_content()

        self.generate_footer_content()

    def setup_config(self):
        self._set_default_format()

    def _set_default_format(self):
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(3, 30)
        self.sheet.set_column('A:A', 13)
        self.sheet.set_column('B:B', 13)
        self.sheet.set_column('C:C', 13)
        self.sheet.set_column('D:D', 18)
        self.sheet.set_column('E:E', 13)
        self.sheet.set_column('F:F', 13)
        self.sheet.set_column('G:G', 13)
        self.sheet.set_column('H:H', 13)
        self.sheet.set_column('I:I', 13)

    def _define_formats(self, workbook):
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        format_config = {
            'font_name': 'Arial',
            'font_size': 10,
            'valign': 'vcenter',
            'text_wrap': True,
        }
        self.format_default = workbook.add_format(format_config)

        format_bold = format_config.copy()
        format_bold.update({
            'bold': True,
        })
        self.format_bold = workbook.add_format(format_bold)

        format_bold_center = format_bold.copy()
        format_bold_center.update({
            'align': 'center',
        })
        self.format_bold_center = workbook.add_format(format_bold_center)

        format_right = format_config.copy()
        format_right.update({
            'align': 'right',
        })
        self.format_right = workbook.add_format(format_right)

        format_center = format_config.copy()
        format_center.update({
            'align': 'center',
        })
        self.format_center = workbook.add_format(format_center)

        # ---------------------------------------------------------------------
        # Company Info
        # ---------------------------------------------------------------------
        format_company_info = format_config.copy()
        format_company_info.update({
            'bold': True,
        })

        self.format_company_info = workbook.add_format(
            format_company_info)

        # ---------------------------------------------------------------------
        # Report Template
        # ---------------------------------------------------------------------
        format_template_title = format_config.copy()
        format_template_title.update({
            'bold': True,
            'align': 'center',
        })
        self.format_template_title = workbook.add_format(format_template_title)

        format_template_desc = format_config.copy()
        format_template_desc.update({
            'italic': True,
            'align': 'center',
        })
        self.format_template_desc = workbook.add_format(
            format_template_desc)

        # ---------------------------------------------------------------------
        # Report Title
        # ---------------------------------------------------------------------
        format_report_title = format_config.copy()
        format_report_title.update({
            'bold': True,
            'align': 'center',
            'font_size': 18,
            'font_name': 'Cabiri',
        })
        self.format_report_title = workbook.add_format(format_report_title)

        format_report_date = format_config.copy()
        format_report_date.update({
            'italic': True,
            'align': 'center',
        })
        self.format_report_date = workbook.add_format(format_report_date)

        format_uom = format_config.copy()
        format_uom.update({
            'italic': True,
            'align': 'center',
        })
        self.format_uom = workbook.add_format(format_uom)
        # ---------------------------------------------------------------------
        # Table format
        # ---------------------------------------------------------------------
        format_table = format_config.copy()
        format_table.update({
            'border': True,
            'num_format': '#,##0',
        })
        self.format_table = workbook.add_format(format_table)

        format_table_center = format_table.copy()
        format_table_center.update({
            'align': 'center',
        })
        self.format_table_center = workbook.add_format(format_table_center)

        format_table_right = format_table.copy()
        format_table_right.update({
            'align': 'right',
        })
        self.format_table_right = workbook.add_format(format_table_right)

        format_table_left = format_table.copy()
        format_table_left.update({
            'align': 'left',
        })
        self.format_table_left = workbook.add_format(format_table_left)

        format_table_bold = format_table.copy()
        format_table_bold.update({
            'bold': True,
        })
        self.format_table_bold = workbook.add_format(format_table_bold)

        format_table_bold_center = format_table_bold.copy()
        format_table_bold_center.update({
            'align': 'center',
        })
        self.format_table_bold_center = workbook.add_format(
            format_table_bold_center)

        format_table_bold_right = format_table_bold.copy()
        format_table_bold_right.update({
            'align': 'right',
        })
        self.format_table_bold_right = workbook.add_format(
            format_table_bold_right)

        format_table_bold_left = format_table_bold.copy()
        format_table_bold_left.update({
            'align': 'left',
        })
        self.format_table_bold_left = workbook.add_format(
            format_table_bold_left)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()
        self.sheet.merge_range(
            'A1:B1', u'Đơn vị báo cáo: %s' % company_info['name'],
            self.format_company_info)

        self.sheet.merge_range(
            'A2:B2', u'Địa chỉ: %s' % company_info['address'],
            self.format_company_info)

        self.sheet.merge_range(
            'A3:B3', u'MST:', self.format_company_info)

    def generate_report_template(self):
        self.sheet.merge_range(
            'E1:I1', u'Mẫu số S08 - DN', self.format_template_title)

        msg = u'(Ban hành theo Thông tư số 200/2014/TT-BTC, Ngày 22/12/2014 của BTC)'
        self.sheet.merge_range(
            'E2:I2', msg, self.format_center)

    def generate_report_title(self):
        self.sheet.merge_range(
            'A4:I4', u'SỔ TIỀN GỬI NGÂN HÀNG', self.format_report_title)

        self.sheet.merge_range(
            'A5:I5', u'Loại tiền gởi: %s' % self.get_account_info(),
            self.format_center)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A6:I6', u'Từ %s đến %s' % (
                date_info['date_from'], date_info['date_to']),
            self.format_template_desc)

    def get_currency_unit(self):
        company = self.env.user.company_id
        return company.currency_id.name

    def get_account_info(self):
        account = self.object.account_id
        return '%s %s' % (account.code, account.name)

    def get_date_info(self):
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {'date_from': date_from, 'date_to': date_to}

    def convert_date_format(self, date_str):
        date = datetime.strptime(date_str, DF)
        return datetime.strftime(date, '%d/%m/%Y')

    def generate_content_header(self):
        self.sheet.merge_range(
            'A8:A9', u'Ngày tháng ghi sổ', self.format_table_bold_center)

        self.sheet.merge_range(
            'B8:C8', u'Chứng từ', self.format_table_bold_center)

        self.sheet.write(
            'B9', u'Số hiệu', self.format_table_bold_center)

        self.sheet.write(
            'C9', u'Ngày tháng', self.format_table_bold_center)

        self.sheet.merge_range(
            'D8:D9', u'Diễn giải', self.format_table_bold_center)

        self.sheet.merge_range(
            'E8:E9', u'Tài khoản đối ứng', self.format_table_bold_center)

        self.sheet.merge_range(
            'F8:H8', u'Số tiền', self.format_table_bold_center)

        self.sheet.write(
            'F9', u'Thu (gởi vào)', self.format_table_bold_center)

        self.sheet.write(
            'G9', u'Chi (rút ra)', self.format_table_bold_center)

        self.sheet.write(
            'H9', u'Tồn', self.format_table_bold_center)

        self.sheet.merge_range(
            'I8:I9', u'Ghi chú', self.format_table_bold_center)

    def generate_main_content(self):
        self.row_position = 10
        self.write_begining_balance()
        start_position = self.row_position

        debit = self.begining_balance['debit_balance'] or 0.0
        credit = self.begining_balance['credit_balance'] or 0.0
        balance = debit - credit

        main_contents = self.prepare_main_content()
        remain_balance = balance
        for line in main_contents:
            remain_balance = int(line.get('debit') or 0) - \
                int(line.get('credit') or 0) + remain_balance
            line['remain'] = remain_balance
            create_date = datetime.strptime(
                line['create_date'], '%Y-%m-%d %H:%M:%S.%f')

            self.sheet.write(
                'A%s' % self.row_position,
                self.convert_date_format(create_date.strftime('%Y-%m-%d')),
                self.format_table_left)

            self.sheet.write(
                'B%s' % self.row_position, line['ref'],
                self.format_table_left)

            self.sheet.write(
                'C%s' % self.row_position,
                self.convert_date_format(line['date']),
                self.format_table_left)

            self.sheet.write(
                'D%s' % self.row_position, line['narration'] or '',
                self.format_table)

            self.sheet.write(
                'E%s' % self.row_position, line['counterpart_account'],
                self.format_table_right)

            self.sheet.write(
                'F%s' % self.row_position, line['debit'] or '',
                self.format_table)

            self.sheet.write(
                'G%s' % self.row_position, line['credit'] or '',
                self.format_table)

            self.sheet.write(
                'H%s' % self.row_position, line['remain'] or '',
                self.format_table)

            self.sheet.write(
                'I%s' % self.row_position, line['description'],
                self.format_table_left)
            self.row_position += 1

        # ---------------------------------------------------------------------
        # Write SUM functions
        # ---------------------------------------------------------------------
        for i in range(0, 9):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.write(
            'D%s' % self.row_position, u'Cộng số phát sinh',
            self.format_table_bold_left)

        end_position = self.row_position - 1

        self.sheet.write_formula(
            'F%s' % self.row_position,
            '=SUM(F%s:F%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'G%s' % self.row_position,
            '=SUM(G%s:G%s)' % (start_position, end_position),
            self.format_table_bold_right)
        self.row_position += 1

        for i in range(0, 9):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.write(
            'D%s' % self.row_position, u'Số dư cuối kỳ',
            self.format_table_bold_left)

        self.sheet.write(
            'H%s' % self.row_position, remain_balance or '',
            self.format_table_bold_right)
        return True

    def generate_footer_content(self):
        self.row_position += 2

        self.sheet.merge_range(
            'F{0}:I{0}'.format(self.row_position),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'Thủ quỹ', self.format_bold_center)

        self.sheet.write(
            'D%s' % self.row_position,
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'F{0}:I{0}'.format(self.row_position),
            u'Giám đốc', self.format_bold_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.write(
            'D%s' % self.row_position,
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'F{0}:I{0}'.format(self.row_position),
            u'(Ký, họ tên, đóng dấu)', self.format_center)

    def write_begining_balance(self):
        self.begining_balance = self.get_begining_balance()
        debit = self.begining_balance['debit_balance'] or 0.0
        credit = self.begining_balance['credit_balance'] or 0.0
        balance = debit - credit

        for i in range(0, 9):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.write(
            'D%s' % self.row_position, u'Số dư đầu kỳ',
            self.format_table_bold_left)
        self.sheet.write(
            'H%s' % self.row_position, balance or 0.0,
            self.format_table_bold_right)
        self.row_position += 1

    def get_begining_balance(self):
        params = self.get_params_for_query()
        SQL = """
            SELECT
                (
                    CASE WHEN (sum(aml.debit)-sum(aml.credit)) > 0
                    THEN (sum(aml.debit)-sum(aml.credit))
                    ELSE 0
                     END
                ) AS debit_balance,
                (
                    CASE WHEN (sum(aml.debit)-sum(aml.credit)) < 0
                    THEN abs(sum(aml.debit)-sum(aml.credit))
                    ELSE 0
                     END
                ) AS credit_balance
            FROM account_move_line aml
            INNER JOIN account_move am on am.id=aml.move_id
            INNER JOIN account_account acc on acc.id=aml.account_id
            WHERE aml.date < '%(date_from)s'
                and am.state in %(state)s
                and acc.internal_type = '%(type)s'
                and acc.id = %(account_id)s
                and aml.journal_id IN %(journal_ids)s
        """ % params
        self.env.cr.execute(SQL)
        data = self.env.cr.dictfetchall() or [0, 0]
        return data[0]

    def prepare_main_content(self):
        params = self.get_params_for_query()
        SQL = """
            SELECT
                aml_dr.create_date AS create_date,
                amv.name AS ref,
                aml_dr.date AS date,

                CASE WHEN  acc.id is null then acc_dr.code
                    else acc_cr.code END as counterpart_account,

                CASE
                    WHEN acc.id is null
                        THEN aml_cr.name
                        ELSE aml_dr.name END AS description,
                amv.narration AS narration,
                CASE
                    WHEN acc.id is not null AND aml_dr.debit > 0
                        THEN aml_dr.debit
                    WHEN acc.id is null     AND aml_cr.debit > 0
                        THEN aml_dr.credit
                END debit,

                CASE
                    WHEN acc.id IS NOT NULL AND aml_dr.credit > 0
                        THEN aml_dr.credit
                    WHEN acc.id IS NULL     AND aml_cr.credit > 0
                        THEN aml_dr.debit
                END credit

            FROM account_move_line aml_dr
            JOIN
                account_move_line aml_cr ON aml_dr.counter_move_id = aml_cr.id
            JOIN
                account_move amv ON amv.id = aml_dr.move_id
            LEFT JOIN
                (   SELECT id
                    FROM account_account
                    WhERE id = %(account_id)s) acc
                        ON acc.id = aml_dr.account_id
            JOIN (SELECT id,code FROM account_account) acc_dr
                ON acc_dr.id = aml_dr.account_id
            JOIN (SELECT id,code FROM account_account) acc_cr
                ON acc_cr.id = aml_cr.account_id

            WHERE (aml_dr.account_id=%(account_id)s
                    OR aml_cr.account_id=%(account_id)s )
                AND aml_dr.date >= '%(date_from)s'
                AND aml_dr.date <= '%(date_to)s'
                AND (aml_dr.journal_id IN %(journal_ids)s
                        OR aml_cr.journal_id IN %(journal_ids)s)
                AND amv.state in %(state)s
            ORDER BY aml_dr.date, amv.name
            """ % params
        self.env.cr.execute(SQL)
        data = self.env.cr.dictfetchall()
        return data

    def get_params_for_query(self):
        params = {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'type': self.object.account_id.internal_type,
            'account_id': self.object.account_id.id,
            'journal_ids': tuple(self.object.journal_ids.ids or [-1, -1])
        }
        if self.object.target_move == 'posted':
            params['state'] = ('posted', '1')
        else:
            params['state'] = ('posted', 'draft')
        return params


CashBankBookReportXlsx(
    'report.cash_bank_book_report_xlsx', 'cash.bank.book.wizard')
