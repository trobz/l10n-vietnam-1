# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from datetime import datetime

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF


class GeneralReceivablePayableLedgerReportXlsx(ReportXlsx):
    '''Generate Receivable Payable Balance Report'''

    def generate_xlsx_report(self, workbook, data, objects):
        '''Generate xlsx report'''
        self.object = objects[0]
        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self.setup_config()

        # generate header
        self.generate_company_info()
        # self.generate_report_template()

        # generate main content
        self.generate_report_title()
        self.generate_content_header()
        self.generate_main_content()

        self.generate_footer_content()

    def setup_config(self):
        '''Setup config'''
        self._set_default_format()

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
            'font_size': 14,
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

    def _set_default_format(self):
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(4, 20)
        self.sheet.set_row(9, 20)
        self.sheet.set_row(10, 20)

        self.sheet.set_column('A:A', 13)
        self.sheet.set_column('B:B', 20)
        self.sheet.set_column('C:C', 15)
        self.sheet.set_column('D:D', 15)
        self.sheet.set_column('E:E', 15)
        self.sheet.set_column('F:F', 15)
        self.sheet.set_column('G:G', 15)
        self.sheet.set_column('H:H', 15)

    def generate_company_info(self):
        '''Generate company info'''
        company_info = self.object.env.user.company_id.get_company_info()

        self.sheet.merge_range(
            'A1:C1', u'Đơn vị báo cáo: ' + company_info['name'],
            self.format_company_info)

        self.sheet.merge_range(
            'A2:C2', u'Địa chỉ: ' + company_info['address'],
            self.format_company_info)

        self.sheet.merge_range(
            'A3:C3', u'MST: ', self.format_company_info)

    def generate_report_title(self):
        '''Generate report title'''
        self.sheet.merge_range(
            'A5:I5', self.get_report_name(), self.format_report_title)

        self.sheet.merge_range(
            'A6:I6', u'Tài khoản: %s' % self.get_account_info(),
            self.format_center)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A7:I7', u'Từ %s đến %s' % (
                date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.write(
            'H9', u'Đơn vị tính: %s' % self.get_currency_unit(),
            self.format_uom)

    def get_currency_unit(self):
        '''Get currency unit'''
        company = self.env.user.company_id
        return company.currency_id.name

    def get_account_info(self):
        '''Get account info'''
        account = self.object.account_id
        return '%s %s' % (account.code, account.name)

    def get_date_info(self):
        '''Get date info'''
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {'date_from': date_from, 'date_to': date_to}

    def convert_date_format(self, date_str):
        '''Conver date format'''
        date = datetime.strptime(date_str, DF)
        return datetime.strftime(date, '%d/%m/%Y')

    def get_report_name(self):
        '''Get report name'''
        if self.object.account_type == 'payable':
            return u'SỔ TỔNG HỢP CÔNG NỢ PHẢI TRẢ'

        return u'SỔ TỔNG HỢP CÔNG NỢ PHẢI THU'

    def generate_content_header(self):
        '''Generate content report header'''
        self.sheet.merge_range(
            'A10:A11', u'Mã KH', self.format_table_bold_center)

        self.sheet.merge_range(
            'B10:B11', u'Tên Khách Hàng', self.format_table_bold_center)

        self.sheet.merge_range(
            'C10:D10', u'Số dư đầu kỳ', self.format_table_bold_center)

        self.sheet.write(
            'C11', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'D11', u'Có', self.format_table_bold_center)

        self.sheet.merge_range(
            'E10:F10', u'Phát sinh trong kỳ', self.format_table_bold_center)

        self.sheet.write(
            'E11', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'F11', u'Có', self.format_table_bold_center)

        self.sheet.merge_range(
            'G10:H10', u'Số dư cuối kỳ', self.format_table_bold_center)

        self.sheet.write(
            'G11', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'H11', u'Có', self.format_table_bold_center)

    def generate_main_content(self):
        '''Generate main content'''
        self.row_position = 12
        start_position = self.row_position

        main_contents = self.prepare_main_content()

        for line in main_contents:
            self.sheet.write(
                'A%s' % self.row_position, line['ref'],
                self.format_table)

            self.sheet.write(
                'B%s' % self.row_position, line['partner_name'],
                self.format_table)

            self.sheet.write(
                'C%s' % self.row_position, line['begin_debit_balance'],
                self.format_table)

            self.sheet.write(
                'D%s' % self.row_position, line['begin_credit_balance'],
                self.format_table_right)

            self.sheet.write(
                'E%s' % self.row_position, line['in_debit'],
                self.format_table)

            self.sheet.write(
                'F%s' % self.row_position, line['in_credit'],
                self.format_table)

            self.sheet.write(
                'G%s' % self.row_position, line['end_debit_balance'],
                self.format_table)

            self.sheet.write(
                'H%s' % self.row_position, line['end_credit_balance'],
                self.format_table)

            self.row_position += 1

        # ---------------------------------------------------------------------
        # Write SUM functions
        # ---------------------------------------------------------------------
        for i in range(0, 8):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position), u'Tổng cộng',
            self.format_table_bold_center)

        end_position = self.row_position - 1

        self.sheet.write_formula(
            'C%s' % self.row_position,
            '=SUM(C%s:C%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'D%s' % self.row_position,
            '=SUM(D%s:D%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'E%s' % self.row_position,
            '=SUM(E%s:E%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'F%s' % self.row_position,
            '=SUM(F%s:F%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'G%s' % self.row_position,
            '=SUM(G%s:G%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'H%s' % self.row_position,
            '=SUM(H%s:H%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.row_position += 1

        for i in range(0, 8):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)

        return True

    def get_payment_term_days(self, days):
        '''Get payment term days'''
        if days and days > 0:
            return u'%s ngày' % days

        return ''

    def generate_footer_content(self):
        '''Generate footer content'''
        self.row_position += 2

        self.sheet.merge_range(
            'G{0}:H{0}'.format(self.row_position),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'Người ghi sổ', self.format_bold_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'G{0}:H{0}'.format(self.row_position),
            u'Giám đốc', self.format_bold_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'G{0}:H{0}'.format(self.row_position),
            u'(Ký, họ tên, đóng dấu)', self.format_center)

    def prepare_main_content(self):
        params = self.get_params_for_query()
        sql = """
            SELECT partner_id,
                ref,
                partner_name,
                (CASE WHEN (sum(begin_debit)-sum(begin_credit)) > 0
                    THEN (sum(begin_debit)-sum(begin_credit))
                    ELSE 0
                END) AS begin_debit_balance,
                (CASE WHEN (sum(begin_debit)-sum(begin_credit)) < 0
                    THEN abs(sum(begin_debit)-sum(begin_credit))
                    ELSE 0
                END) AS begin_credit_balance,

                SUM(in_debit) as in_debit,
                SUM(in_credit) as in_credit,

                (CASE WHEN ((sum(in_debit)-sum(in_credit)) +
                    (sum(begin_debit)-sum(begin_credit))) > 0
                    THEN ((sum(in_debit)-sum(in_credit)) +
                    (sum(begin_debit)-sum(begin_credit)))
                    ELSE 0
                END) AS end_debit_balance,

                (CASE WHEN ((sum(in_debit)-sum(in_credit)) +
                    (sum(begin_debit)-sum(begin_credit))) < 0
                    THEN abs((sum(in_debit)-sum(in_credit)) +
                    (sum(begin_debit)-sum(begin_credit)))
                    ELSE 0
                END) AS end_credit_balance

            FROM (
                SELECT aml.partner_id,
                    rp.ref as ref,
                    rp.name as partner_name,
                    sum(aml.credit) as begin_credit,
                    sum(aml.debit) as begin_debit,
                    0 as in_credit,
                    0 as in_debit

                FROM account_move_line aml
                INNER JOIN account_move am on am.id=aml.move_id
                LEFT JOIN res_partner rp on rp.id=aml.partner_id
                INNER JOIN account_account acc on acc.id=aml.account_id

                WHERE
                    -- Add parameter for account type
                    acc.internal_type='%(type)s'
                    and acc.id = %(account_id)s
                    and aml.date < '%(date_from)s'
                    and am.state in %(state)s
                    and (CASE WHEN %(partner_id)s=-1
                        THEN TRUE ELSE aml.partner_id=%(partner_id)s END)

                GROUP BY %(group_by)s

                UNION ALL

                SELECT aml.partner_id,
                    rp.ref as ref,
                    rp.name as partner_name,
                    0 as begin_credit,
                    0 as begin_debit,
                    sum(aml.credit) as in_credit,
                    sum(aml.debit) as in_debit

                FROM account_move_line aml
                INNER JOIN account_move am on am.id=aml.move_id
                LEFT JOIN res_partner rp on rp.id=aml.partner_id
                INNER JOIN account_account acc on acc.id=aml.account_id

                WHERE
                    -- Add parameter for account type
                    acc.internal_type='%(type)s'
                    and acc.id = %(account_id)s
                    and aml.date between '%(date_from)s' and '%(date_to)s'
                    and am.state in %(state)s
                    and (CASE WHEN %(partner_id)s=-1
                        THEN TRUE ELSE aml.partner_id=%(partner_id)s END)

                GROUP BY %(group_by)s) as A
            GROUP BY %(group_by)s
            ORDER BY 1
        """ % params
        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()
        return data

    def get_params_for_query(self):
        '''Get params for query'''
        journal_ids = self.object.journal_ids.ids + [-1, -1]
        params = {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'type': self.object.account_type,
            'account_id': self.object.account_id.id,
            'journal_ids': tuple(journal_ids),
            'partner_id': self.object.partner_id.id or -1,
            'company_ids': ','.join([str(
                com_id) for com_id in self.env.user.company_ids.ids]),
            'group_by': '1,2,3'
        }

        if self.object.target_move == 'posted':
            params['state'] = ('posted', '1')

        elif self.object.target_move == 'all':
            params['state'] = ('posted', 'draft')

        return params


GeneralReceivablePayableLedgerReportXlsx(
    'report.general_receivable_payable_balance_xlsx',
    'account.payable.receivable.balance')
