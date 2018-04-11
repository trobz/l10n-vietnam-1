# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from openerp.api import Environment
from openerp import SUPERUSER_ID
from datetime import datetime, timedelta
from pytz import timezone
import logging
_logger = logging.getLogger(__name__)


class AccountProfitAndLossXlsx(ReportXlsx):

    def create_xlsx_report(self, ids, data, report):
        self.env = Environment(self.env.cr, SUPERUSER_ID, self.env.context)
        return super(AccountProfitAndLossXlsx, self).create_xlsx_report(
            ids, data, report)

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]

        # choose responsible user of company on wizard
#         user_id = self.object.company_id.resp_user_id.id
#         self.object = self.object.sudo(user=user_id)
#         self.env = Environment(self.env.cr, user_id, self.env.context)

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
            'num_format': '#,##0;[Red](#,##0)',
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
            'num_format': '@',
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

        self.sheet.set_column('A:A', 40)
        self.sheet.set_column('B:B', 10)
        self.sheet.set_column('C:C', 10)
        self.sheet.set_column('D:D', 15)
        self.sheet.set_column('E:E', 15)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()
        info = u'{0}\n{1}\n{2}'.format(
            u'Đơn vị báo cáo: ' + company_info['name'],
            u'Địa chỉ: ' + company_info['address'],
            u'MST: '
        )
        self.sheet.merge_range(
            'A1:B3', info,
            self.format_company_info)

        self.sheet.merge_range(
            'C1:E3', u'Mẫu số S38 – DN\n'
                     u'(Ban hành theo Thông tư số 200/2014/TT-BTC, '
                     u'Ngày 22/12/2014 của BTC)',
            self.format_bold_center)

    def generate_report_title(self):
        self.sheet.merge_range(
            'A{0}:E{0}'.format(4),
            u'BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH',
            self.format_report_title)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A{0}:E{0}'.format(5),
            u'Từ %s đến %s' % (date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(7),
            u'Đơn vị tính: %s' % self.get_currency_unit(),
            self.format_uom)

    def get_currency_unit(self):
        company = self.env.user.company_id
        return company.currency_id.name

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
        self.ws_row = 8
        self.sheet.write('A{}'.format(self.ws_row), u'Chỉ tiêu',
                         self.format_table_bold_center)
        self.sheet.write('B{}'.format(self.ws_row), u'Mã số',
                         self.format_table_bold_center)
        self.sheet.write('C{}'.format(self.ws_row), u'Thuyết minh',
                         self.format_table_bold_center)
        self.sheet.write('D{}'.format(self.ws_row), u'Kỳ báo cáo',
                         self.format_table_bold_center)
        self.sheet.write('E{}'.format(self.ws_row), u'Cùng kỳ năm trước',
                         self.format_table_bold_center)

    # Data region
    def get_params_for_query(self):
        user_timezone = self.object.env.user.tz
        user_timezone = timezone(user_timezone or 'UTC')

        from_date = datetime.strptime(self.object.date_from, DF)
        from_date = user_timezone.localize(from_date)

        to_date = datetime.strptime(self.object.date_to, DF)
        to_date = user_timezone.localize(to_date)
        to_date += timedelta(days=1)

        return {'date_from': from_date,
                'date_to': to_date,
                'last_date_from': from_date - timedelta(days=365),
                'last_date_to': to_date - timedelta(days=365),
                'journal_ids': 'IN {}'.format(tuple(
                    self.object.journal_ids.ids + [-1, -1]))}

    def get_account_data(self, acc_dr_ids, acc_cr_ids, target_move=False):

        if not acc_dr_ids or not acc_cr_ids:
            return 0.0
        # in case of account move line have counter_move_id is null
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        #    cr 33311: 1    counterpart_id
        #    cr 511: 14     counterpart_id
        # =====> sum = credit cr (33311,511)
        params = self.get_params_for_query()
        params.update(
            {'acc_dr_ids': tuple(acc_dr_ids + [-1, -1]),
             'acc_cr_ids': tuple(acc_cr_ids + [-1, -1]),
             'amv_cr_state': target_move == 'posted' and
             ''' AND amv_cr.state = 'posted' ''' or '',
             'amv_dr_state': target_move == 'posted' and
             ''' AND amv_dr.state = 'posted' ''' or ''})
        sql = '''
            SELECT
                COALESCE(SUM(
                    CASE WHEN move_cr.date < '%(date_to)s' AND
                        move_cr.date >= '%(date_from)s'
                        THEN move_cr.credit END),0)
                    AS amount,
                COALESCE(SUM(
                    CASE WHEN move_cr.date < '%(last_date_to)s' AND
                        move_cr.date >= '%(last_date_from)s'
                        THEN move_cr.credit END),0)
                    AS last_amount

            FROM account_move_line move_cr
            LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
            WHERE move_cr.account_id in %(acc_cr_ids)s
                AND move_cr.credit <> 0.0
                AND ((move_cr.date < '%(date_to)s' AND
                move_cr.date >= '%(date_from)s') OR
                (move_cr.date < '%(last_date_to)s' AND
                move_cr.date >= '%(last_date_from)s'))
                %(amv_cr_state)s
                AND move_cr.counter_move_id in (
                    SELECT move_dr.id
                    FROM account_move_line move_dr
                    LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
                    WHERE move_dr.account_id in %(acc_dr_ids)s
                            AND move_dr.counter_move_id is null
                            AND  move_dr.debit <> 0.0
                            AND ((move_dr.date < '%(date_to)s' AND
                                move_dr.date >= '%(date_from)s') OR
                                (move_dr.date < '%(last_date_to)s' AND
                                move_dr.date >= '%(last_date_from)s'))
                             %(amv_dr_state)s
            )
            '''

        # in case of account move line have counter_move_id
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        # dr 515: 1
        #    cr 33311: 16
        # =====> sum = dr debit (111,511)
        sql += '''
            UNION ALL

            SELECT
                COALESCE(SUM(
                    CASE WHEN move_dr.date < '%(date_to)s' AND
                        move_dr.date >= '%(date_from)s'
                        THEN move_dr.debit END),0)
                AS amount,
                COALESCE(SUM(
                    CASE WHEN move_dr.date < '%(last_date_to)s' AND
                    move_dr.date >= '%(last_date_from)s'
                    THEN move_dr.debit END),0)
                AS last_amount
            FROM account_move_line move_dr
            LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
            WHERE move_dr.account_id in %(acc_dr_ids)s
                AND move_dr.debit <> 0.0
                AND ((move_dr.date < '%(date_to)s' AND
                    move_dr.date >= '%(date_from)s') OR
                    (move_dr.date < '%(last_date_to)s' AND
                    move_dr.date >= '%(last_date_from)s'))
                 %(amv_dr_state)s

                AND move_dr.counter_move_id in (
                    SELECT move_cr.id
                    FROM account_move_line move_cr
                    LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
                    WHERE move_cr.account_id in %(acc_cr_ids)s
                            AND move_cr.counter_move_id is null
                            AND  move_cr.credit <> 0.0
                            AND ((move_cr.date < '%(date_to)s' AND
                                move_cr.date >= '%(date_from)s') OR
                                (move_cr.date < '%(last_date_to)s' AND
                                move_cr.date >= '%(last_date_from)s'))
                     %(amv_cr_state)s
            )'''

        sql = sql % params

        self.env.cr.execute(sql)
        result = self.env.cr.dictfetchall()
        current_values = sum([item['amount']
                              for item in result if item['amount']])
        last_values = sum([item['last_amount']
                           for item in result if item['last_amount']])
        return current_values, last_values

    def compute_data(self):
        ret = {}
        account_obj = self.env['account.account']
        profit_loss_objs = self.env['profit.and.loss.config'].search([])

        for profit_loss_obj in profit_loss_objs:
            # for each element of profit and loss
            key = int(profit_loss_obj.code)
            account_ids = [acc.id for acc in profit_loss_obj.account_ids]
            counterpart_account_ids = [
                acc.id for acc in profit_loss_obj.counterpart_account_ids]

            # in case account ids is null, it means getting all account_ids
            all_account_ids = account_obj.search([])
            account_ids = account_ids or all_account_ids.ids
            counterpart_account_ids = \
                counterpart_account_ids or all_account_ids.ids

            # default, is_debit=True : get all total debit

            debit_account_ids = account_ids
            credit_account_ids = counterpart_account_ids

            # in case is_credit=True : get all total credit
            if profit_loss_obj.is_credit:
                debit_account_ids = counterpart_account_ids
                credit_account_ids = account_ids

            target_move = self.object.target_move
            current_values, last_values = self.get_account_data(
                debit_account_ids, credit_account_ids, target_move)

            # in case is_exception: get total_debit - total_credit
            if profit_loss_obj.exception:
                inverted_current_values, inverted_last_values = \
                    self.get_account_data(credit_account_ids,
                                          debit_account_ids, target_move)
                # For account 8211 and 8212
                # if the counterpart account is 911 and 911 is on credit side,
                # value is equal to debit * -1
                # if the counterpart account is 911 and 911 is on debit side,
                # value is equal to credit
                if key in (51, 52):
                    current_values = -1 * current_values \
                        if current_values != 0 else inverted_current_values
                    last_values = -1 * last_values \
                        if last_values != 0 else inverted_last_values
                else:
                    # total_debit - total_credit
                    current_values -= inverted_current_values
                    last_values -= inverted_last_values
            ret.update({key: {'item': profit_loss_obj.item,
                              'code': profit_loss_obj.code,
                              'now': current_values,
                              'last': last_values}})

        return ret

    def _write_line(self, ws_row=1, data=None, f_cell=None):
        data = data or {}
        f_cell = f_cell or self.format_table
        self.sheet.write('A{}'.format(ws_row), data.get('item', ''), f_cell)
        self.sheet.write(
            'B{}'.format(ws_row), data.get('code', ''),
            f_cell == self.format_table and self.format_table_center or
            f_cell == self.format_table_bold and
            self.format_table_bold_center or f_cell
        )
        self.sheet.write('C{}'.format(ws_row), data.get('note', ''), f_cell)
        self.sheet.write('D{}'.format(ws_row), data.get('now', ''), f_cell)
        self.sheet.write('E{}'.format(ws_row), data.get('last', ''), f_cell)

    def generate_main_content(self):
        res = self.compute_data()
        ws_row = 9
        self._write_line(ws_row,
                         dict(item=1, code=2, note=3, now=4, last=5),
                         self.format_table_bold_center)
        ws_row += 1
        self._write_line(ws_row, res[int('01')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('02')], self.format_table)
        ws_row += 1

        code = '10'
        res.update({int(code): dict(
            item=u'3. Doanh thu bán hàng thuần và cung cấp dịch vụ '
                 u'(10 = 01 - 02)',
            code=code, note='',
            now=res[int('01')]['now'] - res[int('02')]['now'],
            last=res[int('01')]['last'] - res[int('02')]['last']
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        self._write_line(ws_row, res[int('11')], self.format_table)
        ws_row += 1

        code = '20'
        res.update({int(code): dict(
            item=u'5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ',
            code=code, note='',
            now=res[int('10')]['now'] - res[int('11')]['now'],
            last=res[int('10')]['last'] - res[int('11')]['last']
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        self._write_line(ws_row, res[int('21')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('22')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('23')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('25')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('26')], self.format_table)
        ws_row += 1

        code = '30'
        res.update({int(code): dict(
            item=u'10. Lợi nhuận thuần từ hoạt động kinh doanh',
            code=code, note='',
            now=res[int('20')]['now'] +
            (res[int('21')]['now'] - res[int('22')]['now']) -
            (res[int('25')]['now'] + res[int('26')]['now']),
            last=res[int('20')]['last'] +
            (res[int('21')]['last'] - res[int('22')]['last']) -
            (res[int('25')]['last'] + res[int('26')]['last'])
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        self._write_line(ws_row, res[int('31')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('32')], self.format_table)
        ws_row += 1

        code = '40'
        res.update({int(code): dict(
            item=u'13. Lợi nhuận khác (40 = 31 -32)',
            code=code, note='',
            now=res[int('31')]['now'] - res[int('32')]['now'],
            last=res[int('31')]['last'] - res[int('32')]['last']
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        code = '50'
        res.update({int(code): dict(
            item=u'14. Tổng lợi nhuận kế toán trước thuế',
            code=code, note='',
            now=res[int('30')]['now'] + res[int('40')]['now'],
            last=res[int('30')]['last'] + res[int('40')]['last']
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        self._write_line(ws_row, res[int('51')], self.format_table)
        ws_row += 1
        self._write_line(ws_row, res[int('52')], self.format_table)
        ws_row += 1

        code = '60'
        res.update({int(code): dict(
            item=u'17. Lợi nhuận sau thuế thu nhập doanh nghiệp',
            code=code, note='',
            now=res[int('50')]['now'] - res[int('51')]['now'] -
            res[int('52')]['now'],
            last=res[int('50')]['last'] - res[int('51')]['last'] -
            res[int('52')]['last']
        )})
        self._write_line(ws_row, res[int(code)], self.format_table_bold)
        ws_row += 1

        code = '70'
        res.update({int(code): dict(
            item=u'18. Lãi cơ bản trên cổ phiếu (*)',
            code=code, note='',
            now='', last=''
        )})
        self._write_line(ws_row, res[int(code)], self.format_table)
        ws_row += 1

        return True

    def generate_footer_content(self):
        self.ws_row = 32

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.ws_row),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.ws_row += 1

        self.sheet.write(
            'A{0}'.format(self.ws_row),
            u'Người lập biểu', self.format_bold_center)

        self.sheet.merge_range(
            'B{0}:C{0}'.format(self.ws_row),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.ws_row),
            u'Giám đốc', self.format_bold_center)
        self.ws_row += 1

        self.sheet.write(
            'A{0}'.format(self.ws_row),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'B{0}:C{0}'.format(self.ws_row),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.ws_row),
            u'(Ký, họ tên, đóng dấu)', self.format_center)
        self.ws_row += 5

        self.sheet.merge_range(
            'A{0}:E{0}'.format(self.ws_row),
            u'Ghi chú: (*) Chỉ tiêu này chỉ áp dụng đối với công ty cổ phần',
            self.format_table
        )


AccountProfitAndLossXlsx(
    'report.account_profit_and_loss_report_xlsx',
    'account.profit.and.loss.report.wizard')
