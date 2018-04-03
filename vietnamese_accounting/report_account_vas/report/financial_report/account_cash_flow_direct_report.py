# -*- coding: utf-8 -*-
##############################################################################
#
#    Copyright 2009-2018 Trobz (<http://trobz.com>).
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not see <http://www.gnu.org/licenses/>.
#
##############################################################################

from openerp.addons.report_base_vn.report import report_base_vn
from datetime import datetime, date
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
import xlwt
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from .. import report_xls_utils
# define mapping between Mã số (in report template) and accounts(to get values)
# see the template to clear

KEY_LINES = {

    1: [{'dr': ('or', [111, 112, 113]),
         'cr': ('or', [5111, 5112, 5113, 131, 3387, 33311])}],

    - 1: [{'dr': ('or', [521, 532, 33311]), 'cr': ('or', [111, 112, 113])}],

    - 2: [{'dr': ('or', [
        152, 153, 1561, 1562, 1567, 1331, 6271, 6272, 6273,
        6277, 6278, 6411, 6412, 6413, 6415, 6417, 6418, 6421,
        6422, 6423, 6425, 6426, 6427, 6428, 331]),
        'cr': ('or', [111, 112, 113])}],

    - 3: [{'dr': ('or', [334]), 'cr': ('or', [111, 112])}],

    - 4: [{'dr': ('or', [635, 335]), 'cr': ('or', [111, 112, 113])}],

    - 5: [{'dr': ('or', [3334]), 'cr': ('or', [111, 112, 113])}],

    6: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [144, 344, 1331])},
        {'number': 6, 'dr': ('or', [111, 112, 113]),
         'cr': ('or', [711, 33311])},
        {'number': '06', 'dr': ('or', [111, 112, 113]),
         'cr': ('or', [711, 33311])}],

    - 7: [{'dr': ('or', [811, 144, 344, 4311, 4312, 4313, 33311, 33312, 3333,
                         3337, 3338]),
           'cr': ('or', [111, 112, 113])}],

    - 21: [{'dr': ('or', [2111, 2112, 2113, 2114, 2115, 2118]),
            'cr': ('or', [111, 112, 113])},
           {'dr': ('or', [212]), 'cr': ('or', [111, 112, 113])},
           {'dr': ('or', [2131, 2132, 2133, 2134, 2135, 2136, 2138]),
            'cr': ('or', [111, 112, 113])},
           {'dr': ('or', [2411, 2412, 2413]), 'cr': ('or', [111, 112, 113])}
           ],

    22: [{'dr': ('or', [111, 112, 113]),
          'cr': ('or', [2111, 2112, 2113, 2114, 2115, 2118, 212,
                        2131, 2132, 2133, 2134, 2135, 2136, 2138, 2411, 2412,
                        2413])},
         {'number': 22, 'dr': ('or', [111, 112, 113]),
          'cr': ('or', [711, 33311])}],

    - 23: [{'dr': ('or', [1281, 1288, 2281, 2282, 2288]),
            'cr': ('or', [111, 112, 113])}],

    24: [{'dr': ('or', [111, 112, 113]),
          'cr': ('or', [1211, 1212, 1281, 1288, 2281, 2282, 2288])},
         {'number': 24, 'dr': ('or', [111, 112, 113]),
          'cr': ('or', [515])}],

    - 25: [{'dr': ('or', [222, 221]),
            'cr': ('or', [111, 112, 113])}],

    26: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [222])}],

    27: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [4211, 4212])},
         {'number': 27, 'dr': ('or', [111, 112, 113]), 'cr': ('or', [515])}],


    31: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [4111, 4112, 4118])}],

    - 32: [{'dr': ('or', [4111, 4112, 4118]), 'cr': ('or', [111, 112, 113])}],

    33: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [311, 341, 342])}],

    - 34: [{'dr': ('or', [311, 341, 342]), 'cr': ('or', [111, 112, 113])}],
    - 35: [{'dr': ('or', [315, 342]), 'cr': ('or', [111, 112, 113])}],
    - 36: [{'dr': ('or', [4211, 4212]), 'cr': ('or', [111, 112, 113])}],
    50: [{'dr': ('or', []), 'cr': ('or', [])}],

    60: [{'start_period': True, 'dr': ('or', [111, 112, 113]),
          'cr': ('or', [1, 2, 3, 4, 5, 6, 7, 8, 9])}],
    - 60: [{'start_period': True, 'dr': ('or', [1, 2, 3, 4, 5, 6, 7, 8, 9]),
            'cr': ('or', [111, 112, 113])}],

    61: [{'dr': ('or', [111, 112, 113]), 'cr': ('or', [4131, 4132])}],
    - 61: [{'dr': ('or', [4131, 4132]), 'cr': ('or', [111, 112, 113])}],

    70: [{'dr': ('or', []), 'cr': ('or', [])}],

}


class Parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'cash_direct_report_xls'
        self.wizard_data = {}
        self.result = {}
        self.date_from_now = ''
        self.date_to_now = ''
        self.date_from_last = ''
        self.date_to_last = ''
        self.localcontext.update({
            'get_currency': self.get_currency,
            'get_local_value': self.get_local_value,
            'get_lines': self.get_lines,
        })
        self.context = context
#         self.compute_result()

    def get_currency(self):

        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        return res.company_id.currency_id.name

    def set_context(self, objects, datas, ids, report_type=None):
        """
        - get all data from wizard and store to self.wizard_data
        - Call compute_result function to compute all result for report
        """

        self.wizard_data = {}
        if datas:
            self.wizard_data['fiscalyear_id'] = 'fiscalyear_id' in \
                datas['form'] and datas['form']['fiscalyear_id'][0] or False
            self.wizard_data['chart_account_id'] = 'chart_account_id' in \
                datas['form'] and datas['form']['chart_account_id'][0] or False
            self.wizard_data['target_move'] = 'target_move' in \
                datas['form'] and datas['form']['target_move'] or ''
            self.wizard_data['type_report'] = 'type_report' in \
                datas['form'] and datas['form']['type_report'] or False
            self.wizard_data['account'] = 'account' in datas['form'] and \
                datas['form']['account'] or False
            self.wizard_data['filter'] = 'filter' in datas['form'] and \
                datas['form']['filter'] or False
#             if datas['form']['filter'] == 'filter_date':
            self.wizard_data['date_from'] = datas['form']['date_from']
            self.wizard_data['date_to'] = datas['form']['date_to']
#             elif datas['form']['filter'] == 'filter_period':
#                 self.wizard_data[
#                     'period_from'] = datas['form']['period_from'][0]
#                 self.wizard_data['period_to'] = datas['form']['period_to'][0]
            self._get_date()
            # compute all line in report
            self.compute_result()
        return super(Parser, self).set_context(
            objects, datas, ids, report_type=report_type)

    def get_local_value(self, key):
        """
        from key, get value of selt.key
        """
        mapping = {
            'date_from': datetime.strptime(
                self.date_from_now,
                DEFAULT_SERVER_DATE_FORMAT).strftime('%d-%m-%Y'),
            'date_to': datetime.strptime(
                self.date_to_now,
                DEFAULT_SERVER_DATE_FORMAT).strftime('%d-%m-%Y')}
        return mapping.get(key, '')

    def _get_date(self):
        #         if self.wizard_data['filter'] == 'filter_period':
        #             period_obj = self.pool.get('account.period')
        #             period_start = period_obj.browse(
        #                 self.cr, self.uid, self.wizard_data['period_from'])
        #             period_end = period_obj.browse(
        #                 self.cr, self.uid, self.wizard_data['period_to'])
        #             self.date_from_now = period_start.date_start
        #             self.date_to_now = period_end.date_stop
        #         elif self.wizard_data['filter'] == 'filter_date':
        self.date_from_now = self.wizard_data['date_from']
        self.date_to_now = self.wizard_data['date_to']
#         else:
#             fiscalyear = self.pool.get('account.fiscalyear').browse(
#                 self.cr, self.uid, self.wizard_data['fiscalyear_id'])
#             self.date_from_now = fiscalyear.date_start
#             self.date_to_now = fiscalyear.date_stop
        # get last year
        date_from_now = datetime.strptime(self.date_from_now,
                                          DEFAULT_SERVER_DATE_FORMAT)
        date_to_now = datetime.strptime(self.date_to_now,
                                        DEFAULT_SERVER_DATE_FORMAT)
        self.date_from_last = date(
            date_from_now.year - 1, date_from_now.month,
            date_from_now.day).strftime(DEFAULT_SERVER_DATE_FORMAT)
        self.date_to_last = date(
            date_to_now.year - 1, date_to_now.month,
            date_to_now.day).strftime(DEFAULT_SERVER_DATE_FORMAT)
        return True

    def run_sql_compute_or(self, acc_dr_ids, acc_cr_ids, date_from, date_to,
                           number_code=False, start_period=False):
        if not acc_dr_ids or not acc_cr_ids or not date_from or not date_to:
            return 0.0
        # in case of account move line have counter_move_id is null
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        #    cr 33311: 1    counterpart_id
        #    cr 511: 14     counterpart_id
        # =====> sum = credit cr (33311,511)
        params = {
            'acc_dr_ids': tuple(acc_dr_ids + [-1, -1]),
            'acc_cr_ids': tuple(acc_cr_ids + [-1, -1]),
            'date_from': date_from, 'date_to': date_to,
            'number_code': number_code,
            'amv_cr_state': self.wizard_data['target_move'] == 'posted' and
            ''' AND amv_cr.state = 'posted' ''' or '',
            'amv_dr_state': self.wizard_data['target_move'] == 'posted' and
            ''' AND amv_dr.state = 'posted' ''' or ''}
        sql = '''
            SELECT SUM(move_cr.credit) as amount
            FROM account_move_line move_cr
            LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
            WHERE move_cr.account_id in %(acc_cr_ids)s
                    AND move_cr.credit <> 0.0
            '''

        if number_code:
            sql = '''
                SELECT SUM(move_cr.credit) as amount
                FROM account_move_line move_cr
                JOIN account_move amv_cr ON move_cr.move_id = amv_cr.id
                JOIN account_model amo ON amv_cr.account_model_id = amo.id
                WHERE   amo.code = '%(number_code)s' '''

        if start_period:
            sql += ''' AND move_cr.date < '%(date_from)s' '''
        else:
            sql += '''  AND move_cr.date <= '%(date_to)s'
                        AND move_cr.date >= '%(date_from)s' '''

        sql += '''
                    %(amv_cr_state)s
                    AND move_cr.counter_move_id in (
                        SELECT move_dr.id
                        FROM account_move_line move_dr
                        LEFT JOIN account_move amv_dr
                            ON amv_dr.id= move_dr.move_id
                        WHERE move_dr.account_id in %(acc_dr_ids)s
                                AND move_dr.counter_move_id is null
                                AND  move_dr.debit <> 0.0 '''
        if start_period:
            sql += ''' AND move_dr.date < '%(date_from)s'
                        %(amv_dr_state)s
                        )
                    '''
        else:
            sql += ''' AND move_dr.date <= '%(date_to)s'
                        AND move_dr.date >= '%(date_from)s'
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

            SELECT sum(move_dr.debit) as amount
            FROM account_move_line move_dr
            LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
            WHERE move_dr.account_id in %(acc_dr_ids)s
                    AND move_dr.debit <> 0.0 '''

        if number_code:
            '''UNION ALL
                SELECT sum(move_dr.debit) as amount
                FROM account_move_line move_dr
                JOIN account_move amv_dr ON move_dr.move_id = amv_dr.id
                JOIN account_model amo ON amv_dr.account_model_id = amo.id
                WHERE   amo.code = '%(number_code)s' '''
        if start_period:
            sql += ''' AND move_dr.date < '%(date_from)s' '''
        else:
            sql += ''' AND move_dr.date <= '%(date_to)s'
                        AND move_dr.date >= '%(date_from)s' '''
        sql += '''
                     %(amv_dr_state)s
                    AND move_dr.counter_move_id in (
                        SELECT move_cr.id
                        FROM account_move_line move_cr
                        LEFT JOIN account_move amv_cr
                            ON amv_cr.id= move_cr.move_id
                        WHERE move_cr.account_id in %(acc_cr_ids)s
                                AND move_cr.counter_move_id is null
                                AND  move_cr.credit <> 0.0 '''
        if start_period:
            sql += ''' AND move_cr.date < '%(date_from)s'
                         %(amv_cr_state)s
                    )
                    '''
        else:
            sql += ''' AND move_cr.date <= '%(date_to)s'
                        AND move_cr.date >= '%(date_from)s'
                         %(amv_cr_state)s
                    )'''

        sql = sql % params

        self.cr.execute(sql)
        values = [item[0] for item in self.cr.fetchall() if item[0]]
        if values:
            return sum(values)
        return 0.0

    def compute_result_or(self, acc_dr_ids, acc_cr_ids, number_code=False,
                          start_period=False):
        """
        Compute result in case of acc_dr and acc_cr are or
        """
        res = {'last': 0.0, 'now': 0.0}
        if not acc_dr_ids or not acc_cr_ids:
            return res
        # in case of current year
        res.update(
            {'now': self.run_sql_compute_or(
                acc_dr_ids, acc_cr_ids, self.date_from_now, self.date_to_now,
                number_code=number_code, start_period=start_period)})
        # in case of last year
        res.update(
            {'last': self.run_sql_compute_or(
                acc_dr_ids, acc_cr_ids, self.date_from_last, self.date_to_last,
                number_code=number_code, start_period=start_period)})
        return res

    def run_sql_compute_and(
            self, acc_dr_ids, acc_cr_ids, dr_and, cr_and,
            date_from, date_to, number_code=False, start_period=False):
        '''
        TODO: this function will be write later
        '''

    def compute_result_and(self, acc_dr_ids, acc_cr_ids, dr_and, cr_and,
                           number_code=False):
        """
        compute data in case of have and in acc_cr, or acc_dr
        """

        res = {'last': 0.0, 'now': 0.0}
        if not acc_dr_ids or not acc_cr_ids:
            return res
        # compute data for current year
        res.update(
            {'now': self.run_sql_compute_and(
                acc_dr_ids, acc_cr_ids, dr_and, cr_and, self.date_from_now,
                self.date_to_now, number_code=number_code)})
        # compute data for last year
        res.update(
            {'last': self.run_sql_compute_and(
                acc_dr_ids, acc_cr_ids, dr_and, cr_and, self.date_from_last,
                self.date_to_last, number_code=number_code)})
        return res

    def search_acc_move_line(
            self, acc_dr, acc_cr, number_code=False, start_period=False):
        """
        @param: acc_dr: tuble (ex: ('or', list debit account))
                acc_cr: tuble (ex: ('or', list credit account))
        """

        account_obj = self.pool.get('account.account')
        acc_dr_ids = []
        acc_cr_ids = []
        # get all account id of list account debit
        for acc in acc_dr[1]:
            acc_dr_ids += account_obj.search(
                self.cr, self.uid, [('code', '=like', '%s%%' % acc)])
        # get all account id of list account credit
        for acc in acc_cr[1]:
            acc_cr_ids += account_obj.search(
                self.cr, self.uid, [('code', '=like', '%s%%' % acc)])
        if acc_cr[0] == 'or' and acc_dr[0] == 'or':
            return self.compute_result_or(
                acc_dr_ids, acc_cr_ids, number_code=number_code,
                start_period=start_period)
        else:
            return self.compute_result_and(
                acc_dr_ids, acc_cr_ids, dr_and=acc_dr[0] == 'and' or False,
                cr_and=acc_cr[0] == 'and' or False, number_code=number_code,
                start_period=start_period)

    def compute_result(self):
        """
        compute all account
        """

        for key, values in KEY_LINES.iteritems():
            self.result[(key, 'last')] = 0
            self.result[(key, 'now')] = 0
            for item in values:
                acc_dr = item.get('dr', ('or', []))
                acc_cr = item.get('cr', ('or', []))
                number_code = item.get('number', False)
                start_period = item.get('start_period', False)
                value = self.search_acc_move_line(
                    acc_dr, acc_cr, number_code, start_period)
                last_value = self.result[(key, 'last')] + \
                    value.get('last', 0.0)
                now_value = self.result[(key, 'now')] + value.get('now', 0.0)
                self.result.update({
                    (key, 'last'): last_value,
                    (key, 'now'): now_value
                })
        return self.result

    def get_lines(self, key, year):
        """
        @param: key: int ---> column Mã số in template report
        """
        result = self.result.get(
            (key, year), 0.0) - self.result.get((-key, year), 0.0)
        return result


class cash_flow_direct_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False,
                 header=True, store=False):
        super(cash_flow_direct_xls, self).__init__(
            name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;',
            'font_Calibri': 'font: name Calibri;'
        })
        _xs = self.xls_styles.copy()
        self.cell_unit_style = xlwt.easyxf(_xs['wrap'] + _xs['right'] +
                                           _xs['top'] + _xs['italic'])
        self.cell_address_style = xlwt.easyxf(_xs['bold'] + _xs['wrap'] +
                                              _xs['left'] + _xs['top'])
        self.cell_decree_style = xlwt.easyxf(_xs['wrap'] +
                                             _xs['center'] + _xs['top'])
        self.cell_title_style = xlwt.easyxf(_xs['bold'] + _xs['wrap'] +
                                            _xs['center'] + _xs['middle'] +
                                            _xs['fontsize_350'] +
                                            _xs['font_Calibri'])
        self.cell_bold_left = xlwt.easyxf(_xs['bold'] + _xs['wrap'] +
                                          _xs['left'] + _xs['borders_all'])
        self.cell_bold_left_italic = xlwt.easyxf(_xs['bold'] +
                                                 _xs['wrap'] +
                                                 _xs['left'] +
                                                 _xs['borders_all'] +
                                                 _xs['italic'])

        self.cell_normal_border_left = xlwt.easyxf(_xs['wrap'] +
                                                   _xs['left'] +
                                                   _xs['borders_all'])
        self.cell_normal_border_right = \
            xlwt.easyxf(_xs['wrap'] +
                        _xs['right'] +
                        _xs['borders_all'],
                        num_format_str='#,##0 ;(#,##0)')

    def generate_xls_header(
            self, _p, _xs, data, objects, wb, ws, row_pos, report_name):

        # Title address 1
        c_specs = [
            ('company_name', 2, 0, 'text',
             u'Đơn vị báo cáo: %s' % _p.get_company()['name'] or '', '',
             self.cell_address_style),
            ('form_serial', 3, 0, 'text', u'Mẫu số B 03 – DN', '',
             self.normal_style_bold)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title address 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 500
        c_specs = [
            ('company_name', 2, 0, 'text',
             u'Địa chỉ: %s' % _p.get_company()['address'] or '', '',
             self.cell_address_style),
            ('form_serial', 3, 0, 'text',
             u'(Ban hành theo QĐ số 15/2006/QĐ-BTC ngày 20/03/2006 của Bộ trưởng BTC)', '',
             self.cell_decree_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [('empty1', 1, 0, 'text', '', None)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2

        c_specs = [('payment_journal', 5, 0, 'text', report_name, None,
                    self.cell_title_style)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title (Theo phương pháp trực tiếp) (*)
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [('amount_on_account', 5, 0, 'text',
                    u'(Theo phương pháp trực tiếp) (*)')]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [(
            'from_to', 5, 0, 'text',
            u'Từ ...%s... đến ...%s...' % (_p.get_local_value('date_from'),
                                           _p.get_local_value('date_to')))]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 1 empty line
        c_specs = [('empty2', 1, 0, 'text', '')]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [('empty3', 3, 0, 'text', ''),
                   ('unit', 2, 0, 'text',
                    u'Đơn vị tính: %s' % _p.get_currency(), '',
                    self.cell_unit_style)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Header Title 1
        row_title_body_pos = row_pos
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 45, 'text', u'Chỉ tiêu', None),
            ('col2', 1, 10, 'text', u'Mã số', None),
            ('col3', 1, 22, 'text', u'Thuyết minh', None),
            ('col4', 1, 22, 'text', u'Năm nay', None),
            ('col5', 1, 22, 'text', u'Năm trước', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_bold_borderall,
            set_column_size=True)
        return row_pos

    def prepare_line_data(self, _p):
        res = {'row01':
               [('A', 1, 0, 'text',
                 u'I. Lưu chuyển tiền từ hoạt động kinh doanh', None,
                 self.cell_bold_left),
                ('B', 1, 0, 'text', '', None, self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None, self.normal_style_borderall),
                ('D', 1, 0, 'text', '', None, self.normal_style_borderall),
                ('E', 1, 0, 'text', '', None, self.normal_style_borderall)],

               'row02':
               [('A', 1, 0, 'text',
                 u'1. Tiền thu từ hoạt động nghiệp vụ, cung cấp dịch vụ và doanh thu khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '01', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.cell_normal_border_left),
                ('D', 1, 0, 'number', _p.get_lines(1, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(1, 'last'), None,
                 self.cell_normal_border_right)],

               'row03':
               [('A', 1, 0, 'text',
                 u'2. Tiền chi trả cho người cung cấp hàng hóa và dịch vụ',
                 None, self.cell_normal_border_left),
                ('B', 1, 0, 'text', '02', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(2, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(2, 'last'), None,
                 self.cell_normal_border_right)],

               'row04':
               [('A', 1, 0, 'text',
                 u'3. Tiền chi trả cho người lao động', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '03', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(3, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(3, 'last'), None,
                 self.cell_normal_border_right)],

               'row05':
               [('A', 1, 0, 'text',
                 u'4. Tiền chi trả lãi vay', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '04', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(4, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(4, 'last'), None,
                 self.cell_normal_border_right)],

               'row06':
               [('A', 1, 0, 'text',
                 u'5. Tiền chi nộp thuế thu nhập doanh nghiệp', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '05', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(5, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(5, 'last'), None,
                 self.cell_normal_border_right)],

               'row07':
               [('A', 1, 0, 'text',
                 u'6. Tiền thu khác từ hoạt động kinh doanh', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '06', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(6, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(6, 'last'), None,
                 self.cell_normal_border_right)],

               'row08':
               [('A', 1, 0, 'text',
                 u'7. Tiền chi khác từ hoạt động kinh doanh', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '07', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(7, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(7, 'last'), None,
                 self.cell_normal_border_right)],

               'row09':
               [('A', 1, 0, 'text',
                 u'Lưu chuyển tiền thuần từ hoạt động kinh doanh', None,
                 self.cell_bold_left_italic),
                ('B', 1, 0, 'text', '20', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', '', 'SUM(D12:D18)',
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', '', 'SUM(E12:E18)',
                 self.cell_normal_border_right)],

               'row10':
               [('A', 1, 0, 'text',
                 u'II. Lưu chuyển tiền từ hoạt động đầu tư', None,
                 self.cell_bold_left),
                ('B', 1, 0, 'text', '', None,  self.cell_normal_border_right),
                ('C', 1, 0, 'text', '', None,  self.cell_normal_border_right),
                ('D', 1, 0, 'text', '', None,  self.cell_normal_border_right),
                ('E', 1, 0, 'text', '', None,  self.cell_normal_border_right)],

               'row11':
               [('A', 1, 0, 'text',
                 u'1. Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác',
                 None, self.cell_normal_border_left),
                ('B', 1, 0, 'text', '21', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(21, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(21, 'last'), None,
                 self.cell_normal_border_right)],

               'row12':
               [('A', 1, 0, 'text',
                 u'2. Tiền thu từ thanh lý, nhượng bán TSCĐ và các tài sản dài hạn khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '22', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(22, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(22, 'last'), None,
                 self.cell_normal_border_right)],

               'row13':
               [('A', 1, 0, 'text',
                 u'3. Tiền chi cho vay, mua các công cụ nợ của đơn vị khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '23', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(23, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(23, 'last'), None,
                 self.cell_normal_border_right)],

               'row14':
               [('A', 1, 0, 'text',
                 u'4. Tiền thu từ thanh lý các khoản đầu tư công cụ nợ của đơn vị khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '24', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(24, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(24, 'last'), None,
                 self.cell_normal_border_right)],

               'row15':
               [('A', 1, 0, 'text',
                 u'5. Tiền chi đầu tư góp vốn vào đơn vị khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '25', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(25, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(25, 'last'), None,
                 self.cell_normal_border_right)],

               'row16':
               [('A', 1, 0, 'text',
                 u'6. Tiền thu hồi đầu tư góp vốn vào đơn vị khác', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '26', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(26, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(26, 'last'), None,
                 self.cell_normal_border_right)],

               'row17':
               [('A', 1, 0, 'text',
                 u'7. Tiền thu lãi cho vay, cổ tức và lợi nhuận được chia',
                 None, self.cell_normal_border_left),
                ('B', 1, 0, 'text', '27', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(27, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(27, 'last'), None,
                 self.cell_normal_border_right)],

               'row18':
               [('A', 1, 0, 'text',
                 u'Lưu chuyển tiền thuần từ hoạt động đầu tư', None,
                 self.cell_bold_left_italic),
                ('B', 1, 0, 'text', '30', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', '', 'SUM(D21:D27)',
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', '', 'SUM(E21:E27)',
                 self.cell_normal_border_right)],

               'row19':
               [('A', 1, 0, 'text',
                 u'III. Lưu chuyển tiền từ hoạt động tài chính', None,
                 self.cell_bold_left),
                ('B', 1, 0, 'text', '', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('E', 1, 0, 'text', '', None,
                 self.normal_style_borderall)],

               'row20':
               [('A', 1, 0, 'text',
                 u'1. Tiền thu từ phát hành cổ phiếu, nhận vốn góp của chủ sở hữu', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '31', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(31, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(31, 'last'), None,
                 self.cell_normal_border_right)],

               'row21':
               [('A', 1, 0, 'text',
                 u'2. Tiền chi trả vốn cho các chủ sở hữu, mua lại cổ phiếu của công ty đã phát hành', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '32', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(32, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(32, 'last'), None,
                 self.cell_normal_border_right)],

               'row22':
               [('A', 1, 0, 'text',
                 u'3. Tiền vay ngắn hạn, dài hạn nhận được', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '33', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(33, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(33, 'last'), None,
                 self.cell_normal_border_right)],

               'row23':
               [('A', 1, 0, 'text',
                 u'4. Tiền chi trả nợ gốc vay', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '34', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(34, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(34, 'last'), None,
                 self.cell_normal_border_right)],

               'row24':
               [('A', 1, 0, 'text', u'5. Tiền chi trả nợ thuê tài chính',
                 None, self.cell_normal_border_left),
                ('B', 1, 0, 'text', '35', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(35, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(35, 'last'), None,
                 self.cell_normal_border_right)],

               'row25':
               [('A', 1, 0, 'text',
                 u'6. Cổ tức, lợi nhuận đã trả cho chủ sở hữu', None,
                 self.cell_normal_border_left),
                ('B', 1, 0, 'text', '36', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(36, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(36, 'last'), None,
                 self.cell_normal_border_right)],

               'row26':
               [('A', 1, 0, 'text',
                 u'Lưu chuyển tiền thuần từ hoạt động tài chính', None,
                 self.cell_bold_left_italic),
                ('B', 1, 0, 'text', '40', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', '', 'SUM(D30:D35)',
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', '', 'SUM(E30:E35)',
                 self.cell_normal_border_right)],

               'row27':
               [('A', 1, 0, 'text', u'''Lưu chuyển tiền thuần trong kỳ
(50 = 20+30+40)''', None, self.cell_bold_left),
                ('B', 1, 0, 'text', '50', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', '', 'SUM(D19,D28,D36)',
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', '', 'SUM(E19,E28,E36)',
                 self.cell_normal_border_right)],

               'row28':
               [('A', 1, 0, 'text', u'Tiền và tương đương tiền đầu kỳ',
                 None, self.cell_bold_left),
                ('B', 1, 0, 'text', '60', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(60, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(60, 'last'), None,
                 self.cell_normal_border_right)],

               'row29':
               [('A', 1, 0, 'text',
                 u'Ảnh hưởng của thay đổi tỷ giá hối đoái quy đổi ngoại tệ',
                 None, self.cell_normal_border_left),
                ('B', 1, 0, 'text', '61', None,
                 self.normal_style_borderall),
                ('C', 1, 0, 'text', '', None,
                 self.normal_style_borderall),
                ('D', 1, 0, 'number', _p.get_lines(61, 'now'), None,
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', _p.get_lines(61, 'last'), None,
                 self.cell_normal_border_right)],

               'row30':
               [('A', 1, 0, 'text',
                 u'''Tiền và tương đương tiền cuối kỳ
(70 = 50+60+61)''', None, self.cell_bold_left),
                ('B', 1, 0, 'text', '70', None,
                 self.normal_style_bold_borderall),
                ('C', 1, 0, 'text', 'VII.34', None,
                 self.normal_style_bold_borderall),
                ('D', 1, 0, 'number', '', 'SUM(D37:D39)',
                 self.cell_normal_border_right),
                ('E', 1, 0, 'number', '', 'SUM(E37:E39)',
                 self.cell_normal_border_right)],
               }
        return res

    def generate_xls_data_line(
            self, _p, _xs, data, objects, wb, ws, row_pos, report_name):
        style_line_normal_center = self.get_cell_style(
            ['center', 'borders_all'])
        c_specs = [('A', 1, 0, 'number', '1', None, None),
                   ('B', 1, 0, 'number', '2', None, None),
                   ('C', 1, 0, 'number', '3', None, None),
                   ('D', 1, 0, 'number', '4', None, None),
                   ('E', 1, 0, 'number', '5', None, None)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=style_line_normal_center)
        line_data = self.prepare_line_data(_p)
        line_key = sorted(line_data.keys())
        for col in line_key:
            # Expand column height for long text
            if col in ['row02', 'row03', 'row11', 'row12', 'row13', 'row14',
                       'row20', 'row21', 'row29', 'row27', 'row30']:
                ws.row(row_pos).height_mismatch = True
                ws.row(row_pos).height = 256 * 2
            c_specs = line_data[col]
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data, row_style=self.normal_style)
        return row_pos

    def generate_xls_footer(
            self, _p, _xs, data, objects, wb, ws, row_pos, report_name):
        # Create an empty lines
        c_specs = [('A', 1, 0, 'text', '', None, None)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        c_specs = [('A', 3, 0, 'text', '', None, None),
                   ('B', 2, 0, 'text', u'Lập, ngày … tháng … năm …')]

        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Footer line 1
        c_specs = [('A', 1, 0, 'text', u'Người lập biểu', None,
                    self.normal_style_bold),
                   ('B', 2, 0, 'text', u'Kế toán trưởng', None,
                    self.normal_style_bold),
                   ('B', 2, 0, 'text', u'Giám đốc', None,
                    self.normal_style_bold)]

        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        c_specs = [('A', 1, 0, 'text', u'(Ký, họ tên)'),
                   ('B', 2, 0, 'text', u'(Ký, họ tên)'),
                   ('B', 2, 0, 'text', u'(Ký, họ tên, đóng dấu)')]

        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        return row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        MAX_ROW = 65500
        count = 1
        report_name = u"BÁO CÁO LƯU CHUYỂN TIỀN TỆ"

        ws = wb.add_sheet(report_name, cell_overwrite_ok=True)
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_num_pages = 1
        ws.fit_height_to_pages = 0
        ws.fit_width_to_pages = 1  # allow to print fit one page
        ws.portrait = True
        ws.show_grid = False
        row_pos = 0
        row_pos = self.generate_xls_header(
            _p, _xs, data, objects, wb, ws, row_pos, report_name)
        row_pos = self.generate_xls_data_line(
            _p, _xs, data, objects, wb, ws, row_pos, report_name)
        row_pos = self.generate_xls_footer(
            _p, _xs, data, objects, wb, ws, row_pos, report_name)


cash_flow_direct_xls(
    'report.account_cash_flow_direct_report_xls',
    'account.move.line', parser=Parser)
