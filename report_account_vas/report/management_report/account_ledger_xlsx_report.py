# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from datetime import datetime, timedelta
from pytz import timezone
import logging
_logger = logging.getLogger(__name__)


class AccountLedgerReportXlsx(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]
        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self.setup_config()

        # generate header
        self.generate_company_info()

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
            'align': 'right',
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

        format_table_date = format_config.copy()
        format_table_date.update({
            'border': True,
            'align': 'right',
            'num_format': 'dd/mm/yyyy'
        })
        self.format_table_date = workbook.add_format(format_table_date)

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
        self.sheet.set_row(2, 20)
        self.sheet.set_row(12, 20)
        self.sheet.set_row(13, 20)

        self.sheet.set_column('A:A', 10)
        self.sheet.set_column('B:B', 13)
        self.sheet.set_column('C:C', 10)
        self.sheet.set_column('D:D', 20)
        self.sheet.set_column('E:E', 10)
        self.sheet.set_column('F:F', 10)
        self.sheet.set_column('G:G', 10)
        self.sheet.set_column('H:H', 10)
        self.sheet.set_column('I:I', 10)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()
        info = u'{0}\n{1}\n{2}'.format(
            u'Đơn vị báo cáo: ' + company_info['name'],
            u'Địa chỉ: ' + company_info['address'],
            u'MST: '
        )
        self.sheet.merge_range(
            'A1:C3', info,
            self.format_company_info)

        self.sheet.merge_range(
            'D1:I3', u'Mẫu số S38 – DN\n'
                     u'(Ban hành theo Thông tư số 200/2014/TT-BTC, Ngày 22/12/2014 của BTC)',
            self.format_bold_center)

        return info

    def generate_report_title(self):
        account_info = self.get_account_info()
        self.sheet.merge_range(
            'A{0}:I{0}'.format(4),
            u'SỔ CHI TIẾT TÀI KHOẢN',
            self.format_report_title)

        self.sheet.merge_range(
            'A{0}:I{0}'.format(5),
            u'Tên tài khoản: %s' % account_info['name'],
            self.format_center)

        self.sheet.merge_range(
            'A{0}:I{0}'.format(6),
            u'Số hiệu tài khoản: %s' % account_info['code'],
            self.format_center)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A{0}:I{0}'.format(7),
            u'Từ %s đến %s' % (date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.merge_range(
            'H{0}:I{0}'.format(8),
            u'Đơn vị tính: %s' % self.get_currency_unit(),
            self.format_uom)

    def get_currency_unit(self):
        company = self.env.user.company_id
        return company.currency_id.name

    def get_account_info(self):
        account = self.object.account_id
        return {
            'code': account.code,
            'name': account.name
        }

    def get_date_info(self):
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {
            'date_from': date_from,
            'date_to': date_to
        }

    def convert_date_format(self, date_str):
        date = datetime.strptime(date_str, DF)
        return datetime.strftime(date, '%d/%m/%Y')

    def generate_content_header(self):
        s_row = self.row_position = 9

        self.sheet.merge_range(
            'A%d:A%d' % (s_row, s_row + 1),
            u'Ngày tháng ghi sổ', self.format_table_bold_center)

        self.sheet.merge_range(
            'B%d:C%d' % (s_row, s_row),
            u'Chứng từ', self.format_table_bold_center)

        self.sheet.write(
            'B%d' % (s_row + 1), u'Số hiệu', self.format_table_bold_center)

        self.sheet.write(
            'C%d' % (s_row + 1), u'Ngày tháng', self.format_table_bold_center)

        self.sheet.merge_range(
            'D%d:D%d' % (s_row, s_row + 1),
            u'Diễn giải', self.format_table_bold_center)

        self.sheet.merge_range(
            'E%d:E%d' % (s_row, s_row + 1),
            u'TK đối ứng', self.format_table_bold_center)

        self.sheet.merge_range(
            'F%d:G%d' % (s_row, s_row),
            u'Số phát sinh', self.format_table_bold_center)

        self.sheet.write(
            'F%d' % (s_row + 1), u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'G%d' % (s_row + 1), u'Có', self.format_table_bold_center)

        self.sheet.merge_range(
            'H%d:I%d' % (s_row, s_row),
            u'Số dư', self.format_table_bold_center)

        self.sheet.write(
            'H%d' % (s_row + 1), u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'I%d' % (s_row + 1), u'Có', self.format_table_bold_center)

    # Data region
    def get_params_for_query(self):
        user_timezone = self.object.env.user.tz
        user_timezone = timezone(user_timezone or 'UTC')

        from_date = datetime.strptime(self.object.date_from, DF)
        from_date = user_timezone.localize(from_date)

        to_date = datetime.strptime(self.object.date_to, DF)
        to_date = user_timezone.localize(to_date)
        to_date += timedelta(days=1)

        params = {
            'account_id': self.object.account_id.id,
            'date_from': from_date,
            'date_to': to_date,
            'journal_ids': 'NOT IN {}'.format(tuple([-1, -1]))
        }

        target_move = ''
        if self.object.target_move == 'posted':
            target_move = "AND amv.state = 'posted'"
        params['target_move'] = target_move

        if self.object.journal_ids:
            params['journal_ids'] = 'IN {}'.format(
                tuple(self.object.journal_ids.ids + [-1, -1]))

        return params

    def get_init(self):
        """
        get the beginning debit/credit amount of current account_id
        """
        params = self.get_params_for_query()
        sql = """
                SELECT sum(COALESCE(aml.debit,0) - COALESCE(aml.credit,0))
                FROM account_move_line aml
                JOIN account_move amv ON aml.move_id = amv.id
                WHERE
                    aml.account_id = %(account_id)s
                    AND aml.date < '%(date_from)s'
                    %(target_move)s
                """
        sql = sql % params
        self.env.cr.execute(sql)

        sum_balance = sum([line[0]
                           for line in self.env.cr.fetchall() if line[0]])
        beginning_debit = sum_balance >= 0 and sum_balance or False
        beginning_credit = sum_balance < 0 and abs(sum_balance) or False

        return {'debit': beginning_debit, 'credit': beginning_credit}

    def get_lines(self):
        params = self.get_params_for_query()

        # init to get beginning debit/credit balance
        self.sum_debit = self.sum_credit = 0.0

        sql = """
            SELECT

                to_char(aml_dr.create_date, 'YYYY-MM-DD') AS create_date,
                aml_dr.date AS date,
                amv.ref AS ref,
                amv.narration AS description,

                CASE WHEN acc.id is null then acc_dr.code else acc_cr.code END as counterpart_account,
                CASE
                    WHEN acc.id is not null AND aml_dr.debit > 0 THEN aml_dr.debit
                    WHEN acc.id is null     AND aml_cr.debit > 0 THEN aml_dr.credit
                END debit,

                CASE
                    WHEN acc.id is not null AND aml_dr.credit > 0 THEN aml_dr.credit
                    WHEN acc.id is null AND aml_cr.credit > 0 THEN aml_dr.debit
                END credit


            FROM account_move_line aml_dr

            JOIN
                account_move_line aml_cr ON aml_dr.counter_move_id = aml_cr.id

            JOIN
                account_move amv ON amv.id = aml_dr.move_id

            LEFT JOIN
                (   SELECT id
                    FROM account_account
                    where id = %(account_id)s) acc ON acc.id = aml_dr.account_id
            JOIN (SELECT id,code from account_account) acc_dr ON acc_dr.id = aml_dr.account_id
            JOIN (SELECT id,code from account_account) acc_cr ON acc_cr.id = aml_cr.account_id


            WHERE (aml_dr.account_id=%(account_id)s or aml_cr.account_id=%(account_id)s )
            and aml_dr.date >= '%(date_from)s' and aml_dr.date < '%(date_to)s'
            %(target_move)s
            ORDER BY aml_dr.date
        """

        sql = sql % params
        self.env.cr.execute(sql)
        res = self.env.cr.dictfetchall()
        return res

    def generate_main_content(self):
        s_row = self.row_position = 11

        # Beginning balance
        beginning_balance = self.get_init()
        self.sheet.write('A%s' % s_row, '', self.format_table)
        self.sheet.write('B%s' % s_row, '', self.format_table)
        self.sheet.write('C%s' % s_row, '', self.format_table)
        self.sheet.write('D%s' % s_row,
                         u'Số dư đầu kỳ', self.format_table_bold)
        self.sheet.write('E%s' % s_row, '', self.format_table)
        self.sheet.write('F%s' % s_row, '', self.format_table)
        self.sheet.write('G%s' % s_row, '', self.format_table)
        self.sheet.write('H%s' % s_row,
                         beginning_balance['debit'] or 0,
                         self.format_table_right)
        self.sheet.write('I%s' % s_row,
                         beginning_balance['credit'] or 0,
                         self.format_table_right)
        s_row += 1

        # Data
        sum_debit_balance = sum_credit_balance = 0
        lines_data = self.get_lines()
        for line in lines_data:
            debit = line.get('debit', 0) or 0
            credit = line.get('credit', 0) or 0

            sum_debit_balance += debit
            sum_credit_balance += credit

            self.sheet.write('A%s' % s_row,
                             line.get('create_date', ''),
                             self.format_table_date)
            self.sheet.write('B%s' % s_row,
                             line.get('ref', ''),
                             self.format_table)
            self.sheet.write('C%s' % s_row,
                             line.get('date', ''),
                             self.format_table_date)
            self.sheet.write('D%s' % s_row,
                             line.get('description', ''),
                             self.format_table)
            self.sheet.write('E%s' % s_row,
                             line.get('counterpart_account', ''),
                             self.format_table)
            self.sheet.write('F%s' % s_row,
                             debit,
                             self.format_table_right)
            self.sheet.write('G%s' % s_row,
                             credit,
                             self.format_table_right)
            formula = 'H{0} - I{0} + F{1} - G{1}'.format(s_row - 1, s_row)
            self.sheet.write('H%s' % s_row,
                             '=IF({0} > 0, {0}, 0)'.format(formula),
                             self.format_table_right)
            self.sheet.write('I%s' % s_row,
                             '=IF({0} < 0, ABS({0}), 0)'.format(formula),
                             self.format_table_right)
            s_row += 1

        # Cộng phát sinh
        self.sheet.write('A%s' % s_row, '', self.format_table)
        self.sheet.write('B%s' % s_row, '', self.format_table)
        self.sheet.write('C%s' % s_row, '', self.format_table)
        self.sheet.write('D%s' % s_row,
                         u'Cộng phát sinh', self.format_table_bold)
        self.sheet.write('E%s' % s_row, '', self.format_table)
        self.sheet.write('F%s' % s_row,
                         sum_debit_balance > 0 and sum_debit_balance or '',
                         self.format_table_right)
        self.sheet.write('G%s' % s_row,
                         sum_credit_balance > 0 and sum_credit_balance or '',
                         self.format_table_right)
        self.sheet.write('H%s' % s_row, '', self.format_table)
        self.sheet.write('I%s' % s_row, '', self.format_table)
        s_row += 1

        # Header số dư cuối kỳ
        total_balance = \
            (beginning_balance.get('debit', 0) + sum_debit_balance) - \
            (beginning_balance.get('credit', 0) + sum_credit_balance)
        self.sheet.write('A%s' % s_row, '', self.format_table)
        self.sheet.write('B%s' % s_row, '', self.format_table)
        self.sheet.write('C%s' % s_row, '', self.format_table)
        self.sheet.write('D%s' % s_row,
                         u'Số dư cuối kỳ', self.format_table_bold)
        self.sheet.write('E%s' % s_row, '', self.format_table)
        self.sheet.write('F%s' % s_row, '', self.format_table)
        self.sheet.write('G%s' % s_row, '', self.format_table)
        self.sheet.write('H%s' % s_row,
                         total_balance > 0 and total_balance or '',
                         self.format_table_right)
        self.sheet.write('I%s' % s_row,
                         total_balance < 0 and abs(total_balance) or '',
                         self.format_table_right)
        s_row += 1

        self.row_position = s_row
        return True

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


AccountLedgerReportXlsx(
    'report.account_ledger_report_xlsx', 'account.ledger.wizard')
