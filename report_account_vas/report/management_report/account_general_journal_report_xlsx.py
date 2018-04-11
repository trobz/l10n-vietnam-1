# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from datetime import datetime

from openerp import fields
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT as DTF


class AccountGeneralJournalReportXlsx(ReportXlsx):

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

        # ---------------------------------------------------------------------
        # Table format with background
        # ---------------------------------------------------------------------

        format_table_bold_bg = format_table_bold.copy()
        format_table_bold_bg.update({
            'bg_color': '#FFFE82'
        })
        self.format_table_bold_bg = workbook.add_format(
            format_table_bold_bg)

        format_table_bold_center_bg = format_table_bold_center.copy()
        format_table_bold_center_bg.update({
            'bg_color': '#FFFE82'
        })
        self.format_table_bold_center_bg = workbook.add_format(
            format_table_bold_center_bg)

        format_table_bold_right_bg = format_table_bold_right.copy()
        format_table_bold_right_bg.update({
            'bg_color': '#FFFE82'
        })
        self.format_table_bold_right_bg = workbook.add_format(
            format_table_bold_right_bg)

    def _set_default_format(self):
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(2, 20)
        self.sheet.set_row(4, 20)  # title
        self.sheet.set_row(7, 25)  # header
        self.sheet.set_row(8, 25)  # header

        self.sheet.set_column('A:A', 15)
        self.sheet.set_column('B:B', 20)
        self.sheet.set_column('C:C', 13)
        self.sheet.set_column('D:D', 30)
        self.sheet.set_column('E:E', 13)
        self.sheet.set_column('F:F', 20)
        self.sheet.set_column('G:G', 20)

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
            'E1:G1', u'Mẫu số S03a – DN', self.format_template_title)

        msg = u'(Ban hành theo Thông tư số 200/2014/TT-BTC, ' + \
            u'Ngày 22/12/2014 của BTC)'

        self.sheet.merge_range(
            'E2:G2', msg, self.format_template_desc)

    def generate_report_title(self):
        self.sheet.merge_range(
            'A5:G5', self.get_report_name(), self.format_report_title)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A6:G6', u'Từ %s đến %s' % (
                date_info['date_from'], date_info['date_to']),
            self.format_center)

    def get_date_info(self):
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {
            'date_from': date_from,
            'date_to': date_to
        }

    def convert_date_format(self, date_str, ctx_timestamp=False):
        rec = self.object
        try:
            rec = self.object
            fm = DF
            if ctx_timestamp:
                fm = DTF
            val = datetime.strptime(date_str, fm)
            if ctx_timestamp:
                val = fields.Datetime.context_timestamp(rec, val)
            return datetime.strftime(val, '%d/%m/%Y')

        except:
            return ''

    def get_report_name(self):
        return u'SỔ NHẬT KÝ CHUNG'

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
            'E8:E9', u'Tài khoản', self.format_table_bold_center)

        self.sheet.merge_range(
            'F8:G8', u'Số phát sinh', self.format_table_bold_center)

        self.sheet.write(
            'F9', u'Nợ', self.format_table_bold_center)

        self.sheet.write(
            'G9', u'Có', self.format_table_bold_center)

    def generate_main_content(self):
        self.row_position = 10
        main_contents = self.prepare_main_content()

        account_move_code = 1
        start_position = self.row_position
        for line in main_contents:
            if line['code'] == account_move_code:
                self.generate_summary_line(line)
            else:
                self.generate_detail_line(line)

            self.row_position += 1

        # ---------------------------------------------------------------------
        # Write SUM functions
        # ---------------------------------------------------------------------
        for i in range(0, 7):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)

        self.sheet.write(
            'D%s' % self.row_position, u'Tổng cộng',
            self.format_table_bold_center)

        self.sheet.write_formula(
            'F%s' % self.row_position,
            '=SUM(F%s:F%s)' % (start_position, self.row_position - 1),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'G%s' % self.row_position,
            '=SUM(G%s:G%s)' % (start_position, self.row_position - 1),
            self.format_table_bold_right)

        self.row_position += 1
        return True

    def generate_summary_line(self, line):
        self.sheet.write(
            'A%s' % self.row_position, line['create_date'],
            self.format_table_bold_bg)

        self.sheet.write(
            'B%s' % self.row_position, line['ref'],
            self.format_table_bold_bg)

        self.sheet.write(
            'C%s' % self.row_position, line['date'],
            self.format_table_bold_bg)

        self.sheet.write(
            'D%s' % self.row_position, line['narration'],
            self.format_table_bold_bg)

        self.sheet.write(
            'E%s' % self.row_position, line['account'],
            self.format_table_bold_center_bg)

        self.sheet.write(
            'F%s' % self.row_position, line['debit'] or '',
            self.format_table_bold_right_bg)

        self.sheet.write(
            'G%s' % self.row_position, line['credit'] or '',
            self.format_table_bold_right_bg)

    def generate_detail_line(self, line):
        self.sheet.write(
            'A%s' % self.row_position, '',
            self.format_table)

        self.sheet.write(
            'B%s' % self.row_position, line['ref'],
            self.format_table)

        self.sheet.write(
            'C%s' % self.row_position, line['date'],
            self.format_table)

        self.sheet.write(
            'D%s' % self.row_position, line['narration'],
            self.format_table)

        self.sheet.write(
            'E%s' % self.row_position, line['account'],
            self.format_table_center)

        self.sheet.write(
            'F%s' % self.row_position, line['debit'] or '',
            self.format_table_right)

        self.sheet.write(
            'G%s' % self.row_position, line['credit'] or '',
            self.format_table_right)

    def generate_footer_content(self):
        self.row_position += 2

        self.sheet.write(
            'G%s' % (self.row_position),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'Người lập biểu', self.format_bold_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.write(
            'G%s' % (self.row_position),
            u'Giám đốc', self.format_bold_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.write(
            'G%s' % (self.row_position),
            u'(Ký, họ tên, đóng dấu)', self.format_center)

    def prepare_main_content(self):
        params = self.get_params_for_query()
        sql = """
            SELECT  code, 
                    to_char(create_date + INTERVAL '7 hour', 'DD/MM/YYYY')
                        AS create_date,
                    ref,
                    date,
                    narration,
                    account,
                    debit,
                    credit
            FROM
            (
                SELECT id AS move_id,
                    1 AS code,
                    create_date,
                    name AS ref,
                    to_char(date, 'DD/MM/YYYY') AS date,
                    narration,
                    NULL::character as account,
                    NULL::numeric as debit,
                    NULL::numeric as credit,
                    create_date:: date AS date_order
                    
                FROM account_move
                WHERE date BETWEEN '%(date_from)s' AND '%(date_to)s'
                    AND company_id = %(company_id)s
                    AND state IN %(state)s
            
                UNION ALL
            
                SELECT move_id,
                    2 AS code,
                    NULL as create_date,
                    NULL as ref,
                    NULL as date,
                    NULL AS narration,
                    account.code as account,
                    debit,
                    credit,
                    am.create_date:: date AS date_order
                
                FROM account_move_line aml
                    JOIN account_move am
                        ON am.id = aml.move_id
                    JOIN account_account account
                        ON account.id = aml.account_id
                    WHERE am.date BETWEEN '%(date_from)s' AND '%(date_to)s'
                        AND am.company_id = %(company_id)s
                        AND am.state IN %(state)s
            ) tab
            
            ORDER BY date_order, move_id, code, debit            
        """ % params

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()
        return data

    def get_params_for_query(self):

        journal_ids = self.object.journal_ids.ids
        if journal_ids:
            journal_ids += [-1, -1]
        else:
            journal_ids = [-1, -1]

        params = {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'journal_ids': tuple(journal_ids),
            'company_id': self.object.company_id.id,
        }

        if self.object.target_move == 'posted':
            params['state'] = ('posted', '1')

        elif self.object.target_move == 'all':
            params['state'] = ('posted', 'draft')

        return params


AccountGeneralJournalReportXlsx(
    'report.vas_account_general_journal_xlsx', 'vas.account.general.journal')
