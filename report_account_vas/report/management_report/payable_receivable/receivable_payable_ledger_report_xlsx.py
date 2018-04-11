# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from datetime import datetime

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF


class ReceivablePayableLedgerReportXlsx(ReportXlsx):

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
        self.sheet.set_row(12, 20)

        self.sheet.set_column('A:A', 13)
        self.sheet.set_column('B:B', 13)
        self.sheet.set_column('C:C', 20)
        self.sheet.set_column('D:D', 10)
        self.sheet.set_column('E:E', 10)
        self.sheet.set_column('F:F', 13)
        self.sheet.set_column('G:G', 13)
        self.sheet.set_column('H:H', 13)
        self.sheet.set_column('I:I', 13)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()

        self.sheet.merge_range(
            'A1:C1', u'Đơn vị báo cáo: ' + company_info['name'],
            self.format_company_info)

        self.sheet.merge_range(
            'A2:C2', u'Địa chỉ: ' + company_info['address'],
            self.format_company_info)

        self.sheet.merge_range(
            'A3:C3', u'MST: ', self.format_company_info)

    def generate_report_template(self):
        self.sheet.merge_range(
            'E1:I1', u'Mẫu số S31 - DN', self.format_template_title)

        msg = u'(Ban hành theo Thông tư số 200/2014/TT-BTC ' + \
            u'Ngày 22/12/2014 của Bộ Tài chính)'
        self.sheet.merge_range(
            'E2:I2', msg, self.format_template_desc)

    def generate_report_title(self):
        partner_info = self.get_partner_info()
        self.sheet.merge_range(
            'A5:I5', self.get_report_name(), self.format_report_title)

        self.sheet.merge_range(
            'A6:I6', u'%s: %s' % (
                partner_info['ref_text'], partner_info['ref']),
            self.format_center)

        self.sheet.merge_range(
            'A7:I7', u'%s: %s' % (
                partner_info['name_text'], partner_info['name']),
            self.format_center)

        self.sheet.merge_range(
            'A8:I8', u'Tài khoản: %s' % self.get_account_info(),
            self.format_center)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A9:I9', u'Từ %s đến %s' % (
                date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.write(
            'I10', u'Đơn vị tính: %s' % self.get_currency_unit(),
            self.format_uom)

    def get_currency_unit(self):
        company = self.env.user.company_id
        return company.currency_id.name

    def get_partner_info(self):
        partner = self.object.partner_id
        if self.object.account_type == 'payable':
            ref_text = u'Mã nhà cung cấp'
            name_text = u'Tên nhà cung cấp'

        else:
            ref_text = u'Mã khách hàng'
            name_text = u'Tên khách hàng'

        return {
            'ref_text': ref_text,
            'ref': partner and partner.ref or u'Tất cả',
            'name_text': name_text,
            'name': partner and partner.name or u'Tất cả',
        }

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

    def get_report_name(self):
        if self.object.account_type == 'payable':
            return u'SỔ CHI TIẾT CÔNG NỢ PHẢI TRẢ'

        return u'SỔ CHI TIẾT CÔNG NỢ PHẢI THU'

    def generate_content_header(self):
        self.sheet.merge_range(
            'A11:B11', u'Chứng từ', self.format_table_bold_center)

        self.sheet.write(
            'A12', u'Ngày tháng', self.format_table_bold_center)

        self.sheet.write(
            'B12', u'Số hiệu', self.format_table_bold_center)

        self.sheet.merge_range(
            'C11:C12', u'Diễn giải', self.format_table_bold_center)

        self.sheet.merge_range(
            'D11:D12', u'TK đối ứng', self.format_table_bold_center)

        self.sheet.merge_range(
            'E11:E12', u'Thời hạn được chiết khấu',
            self.format_table_bold_center)

        self.sheet.merge_range(
            'F11:G11', u'Số phát sinh', self.format_table_bold_center)

        self.sheet.write(
            'F12', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'G12', u'Có', self.format_table_bold_center)

        self.sheet.merge_range(
            'H11:I11', u'Số dư', self.format_table_bold_center)

        self.sheet.write(
            'H12', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'I12', u'Có', self.format_table_bold_center)

        self.sheet.merge_range(
            'A13:D13', u'Số dư đầu kỳ', self.format_table_bold_center
        )

        for col_pos in range(4, 9):
            self.sheet.write(12, col_pos, '', self.format_table)

    def generate_main_content(self):
        self.row_position = 13
        start_position = self.row_position
        self.write_begining_balance()

        main_contents = self.prepare_main_content()

        for line in main_contents:
            self.sheet.write(
                'A%s' % self.row_position, line['date'],
                self.format_table)

            self.sheet.write(
                'B%s' % self.row_position, line['serial'],
                self.format_table)

            self.sheet.write(
                'C%s' % self.row_position, line['description'],
                self.format_table)

            self.sheet.write(
                'D%s' % self.row_position, line['counterpart_account'],
                self.format_table_right)

            self.sheet.write(
                'E%s' % self.row_position,
                self.get_payment_term_days(line['payment_term_days']),
                self.format_table)

            self.sheet.write(
                'F%s' % self.row_position, line['debit'] or '',
                self.format_table)

            self.sheet.write(
                'G%s' % self.row_position, line['credit'] or '',
                self.format_table)

            self.sheet.write(
                'H%s' % self.row_position, line['debit_balance'] or '',
                self.format_table)

            self.sheet.write(
                'I%s' % self.row_position, line['credit_balance'] or '',
                self.format_table)

            self.row_position += 1

        # ---------------------------------------------------------------------
        # Write SUM functions
        # ---------------------------------------------------------------------
        for i in range(0, 9):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.merge_range(
            'A{0}:E{0}'.format(self.row_position), u'Tổng cộng',
            self.format_table_bold_center)

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
        self.sheet.merge_range(
            'A{0}:E{0}'.format(self.row_position), u'Số dư cuối kỳ',
            self.format_table_bold_center)

        self.sheet.write_formula(
            'H%s' % self.row_position,
            '=H%s' % (self.row_position - 2),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'I%s' % self.row_position,
            '=I%s' % (self.row_position - 2),
            self.format_table_bold_right)

        return True

    def get_payment_term_days(self, days):
        if days and days > 0:
            return u'%s ngày' % days

        return ''

    def generate_footer_content(self):
        self.row_position += 2

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'Người ghi sổ', self.format_bold_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'Giám đốc', self.format_bold_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'(Ký, họ tên, đóng dấu)', self.format_center)

    def write_begining_balance(self):
        self.begining_balance = self.get_begining_balance()

        debit = self.begining_balance['debit_balance'] or ''
        self.sheet.write(
            'H%s' % self.row_position, debit, self.format_table_right)

        credit = self.begining_balance['credit_balance'] or ''
        self.sheet.write(
            'I%s' % self.row_position, credit, self.format_table_right)
        self.row_position += 1

    def get_begining_balance(self):
        params = self.get_params_for_query()

        sql = """
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
                %(state_condition)s
                and acc.internal_type = '%(type)s'
                and acc.id = %(account_id)s
                %(partner_condition)s
                %(journal_condition)s
        """ % params

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall() or [0, 0]

        return data[0]

    def prepare_main_content(self):
        params = self.get_params_for_query()
        sql = """
            SELECT
                am.name AS serial,
                aml.date AS date,
                am.narration AS description,
                acc_du.code AS counterpart_account,
                COALESCE(aml_counterpart.credit, 0) AS debit,
                COALESCE(aml_counterpart.debit, 0) AS credit,
                aptl.days AS payment_term_days

            FROM account_move_line  aml
                INNER JOIN account_move am
                    ON am.id=aml.move_id

                INNER JOIN account_account acc
                    ON acc.id=aml.account_id

                LEFT JOIN (
                    SELECT *
                    FROM account_move_line
                ) aml_counterpart
                    ON aml_counterpart.id = aml.counter_move_id
                        OR aml_counterpart.counter_move_id = aml.id

                LEFT JOIN account_account acc_du
                    ON acc_du.id=aml_counterpart.account_id

                LEFT JOIN account_invoice ai
                    ON ai.id = aml.invoice_id
                LEFT JOIN account_payment_term apt
                    ON ai.payment_term_id = apt.id
                LEFT JOIN LATERAL (
                    SELECT *
                    FROM account_payment_term_line
                    WHERE payment_id = apt.id
                    ORDER BY sequence, create_date
                    LIMIT 1
                ) aptl ON TRUE

            WHERE aml.date between '%(date_from)s' AND '%(date_to)s'
                %(state_condition)s
                AND acc.internal_type='%(type)s'
                AND acc.id = %(account_id)s
                %(partner_condition)s
                %(journal_condition)s

            ORDER BY date, aml.id, serial, description

        """ % params

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()
        data = self.modify_main_data(data)
        return data

    def modify_main_data(self, data):
        begining_balance = self.begining_balance

        first_line = True
        for index, line in enumerate(data):
            if first_line:
                first_line = False

                balance = begining_balance['debit_balance'] - \
                    begining_balance['credit_balance'] + \
                    line['debit'] - line['credit']

                data[index].update({
                    'debit_balance': balance > 0 and balance or 0,
                    'credit_balance': balance < 0 and abs(balance) or 0,
                })
                continue

            balance = data[index - 1]['debit_balance'] - \
                data[index - 1]['credit_balance'] + \
                line['debit'] - line['credit']

            data[index].update({
                'debit_balance': balance > 0 and balance or 0,
                'credit_balance': balance < 0 and abs(balance) or 0,
            })

        return data

    def get_params_for_query(self):
        journal_ids = self.object.journal_ids.ids
        journal_condition = ''
        if journal_ids:
            journal_ids += [-1, -1]
            journal_condition = "AND aml.journal_id IN %s" % tuple(
                journal_ids)

        partner_id = self.object.partner_id.id
        params = {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'type': self.object.account_type,
            'partner_condition': partner_id and
            "AND aml.partner_id=%s" % partner_id or '',
            'account_id': self.object.account_id.id,
            'journal_condition': journal_condition,
            'extra_condition': ''
        }

        target_move = self.object.target_move
        params['state_condition'] = target_move == 'posted' and \
            "AND am.state = 'posted'" or ''

        return params


ReceivablePayableLedgerReportXlsx(
    'report.general_detail_receivable_payable_balance',
    'account.detailed.payable.receivable.balance')
