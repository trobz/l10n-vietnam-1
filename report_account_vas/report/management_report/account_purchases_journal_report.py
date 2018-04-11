# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import xlwt
from datetime import datetime
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from openerp.addons.report_base_vn.report import report_base_vn
from .. import report_xls_utils
import logging
_logger = logging.getLogger(__name__)


class account_purchases_journal_xls_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(account_purchases_journal_xls_parser, self).__init__(
            cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_account_info': self.get_account_info,
            'is_purchases_journal': True,
        })

    def get_company(self):
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        address_list = [res.company_id.street or '',
                        res.company_id.street2 or '',
                        res.company_id.city or '',
                        res.company_id.state_id and
                        res.company_id.state_id.name or '',
                        res.company_id.country_id and
                        res.company_id.country_id.name or '',
                        ]
        address_list = filter(None, address_list)
        address = ', '.join(address_list)
        vat = res.company_id.vat or ''
        return {'name': name, 'address': address, 'vat': vat}

    def get_wizard_data(self):
        result = {}
        datas = self.localcontext['data']
        if datas:
            result['fiscalyear'] = datas['form'] and datas[
                'form']['fiscalyear_id'] or False
            result['target_move'] = datas['form'] and datas[
                'form']['target_move'] or False
            result['account'] = datas['form'] and datas['form'][
                'account'] and datas['form']['account'][0] or False
            result['filter'] = 'filter' in datas[
                'form'] and datas['form']['filter'] or False
            if datas['form']['filter'] == 'filter_date':
                result['date_from'] = datas['form']['date_from']
                result['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                result['period_from'] = datas['form']['period_from']
                result['period_to'] = datas['form']['period_to']
        return result

    def get_account_info(self):
        account_id = self.get_wizard_data().get('account', '') or False
        if account_id:
            res = self.pool.get('account.account').read(
                self.cr, self.uid, account_id, ['code', 'name'])
            return res
        return {'code': '', 'name': ''}

    def get_lines_data(self):
        """
        Get all account data
        debit_account is 11*
            credit_account
        """
        date_info = self.get_date()
        params = {'target_move': '',
                  'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date']}

        SQL = """

        /* Case 1: counterpart_id is set on Cr

            + one-one
                Dr    |    Cr
                1111  |
                      |    131    counterpart_id
        */

        SELECT  move_cr.date_created,
            amv.name as serial,
            move_cr.date as effective_date,
            move_cr.name as description,
            dr_account.code as debit_code,
            cr_account.code as credit_code,
            move_dr.debit as amount
        FROM
            account_move_line move_dr
            LEFT JOIN account_move amv ON amv.id= move_dr.move_id
            JOIN account_move_line move_cr
                ON move_cr.counter_move_id = move_dr.id
            JOIN account_account dr_account
                ON move_dr.account_id = dr_account.id
            JOIN account_account cr_account
                ON move_cr.account_id = cr_account.id
        WHERE
            move_cr.credit != 0
            AND move_dr.debit != 0
            AND move_dr.date >= '%(date_start)s' AND
                move_dr.date <= '%(date_end)s'
            AND amv.inventory_valuation_picking_type = 'purchase'
            %(target_move)s

        """

        if self.get_wizard_data()['target_move'] == 'posted':
            params['target_move'] = "AND amv.state = 'posted'"

        # order by date_created
        SQL = SQL + " ORDER BY date_created"

#         print "######################### SQL params", SQL % params
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data


class account_purchases_journal_xls(report_xls_utils.generic_report_xls_base):

    def __init__(
            self, name, table, rml=False,
            parser=False, header=True, store=False):
        super(account_purchases_journal_xls, self).__init__(
            name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E', 'G', 'H']
        self.col_specs_template = {
            'A': {
                'lines': [1, 12, _render("line.get('date_created','') and 'date' or 'text'"), _render("datetime.strptime(line.get('date_created','')[:10],'%Y-%m-%d')"), None, self.style_date_right],
                'totals': [1, 0, 'text', None]},

            'B': {
                'lines': [1, 17, 'text', _render("line.get('serial','') or ''"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'C': {
                'lines': [1, 17, _render("line.get('effective_date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('effective_date',None)[:10],'%Y-%m-%d')"), None, self.style_date_right],
                'totals': [1, 0, 'text', None]},

            'D': {
                'lines': [1, 50, 'text', _render("line.get('description','') or ''"), None, self.normal_style_left_borderall],
                'totals': [1, 16, 'text', 'Tổng Cộng', None, self.normal_style_bold_borderall]},

            'E': {
                'lines': [1, 15, 'text', _render("line.get('debit_code',None)"), None, self.normal_style_right_borderall],
                'totals': [1, 0, 'text', None]},

            'G': {
                'lines': [1, 15, 'text', _render("line.get('credit_code',None)"), None, self.normal_style_right_borderall],
                'totals': [1, 0, 'text', None]},

            'H': {
                'lines': [1, 28, 'number', _render("line.get('amount',0)"), None, self.style_decimal],
                'totals': [1, 0, 'number', _render('sum_debit_amount'), None,  self.style_decimal_bold],
                'totals_no_lines': [1, 0, 'number', None, None, self.style_decimal_bold]},
        }

    def generate_xls_header(
            self, _p, _xs, data, objects, wb, ws, row_pos, report_name):
        """

        """

        cell_address_style = self.get_cell_style(['bold', 'wrap', 'left'])
        # Title address 1
        c_specs = [
            ('company_name', 7, 0, 'text', u'Đơn vị: %s' %
             _p.get_company()['name'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title 2
        c_specs = [
            ('company_name', 7, 0, 'text', u'Địa chỉ: %s' %
             _p.get_company()['address'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('company_name', 7, 0, 'text', u'MST: %s' %
             _p.get_company()['vat'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 17, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 50, 'text', '', None),
            ('col5', 1, 15, 'text', '', None),
            ('col6', 1, 15, 'text', '', None),
            ('col7', 1, 22, 'text', '', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "SỔ NHẬT KÝ MUA HÀNG"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_title_style = self.get_cell_style(
            ['bold', 'wrap', 'center', 'middle', 'fontsize_350'])

        c_specs = [
            ('payment_journal', 7, 0, 'text',
                report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('from_to', 7, 0, 'text', 'Từ %s đến %s' % (_p.get_date().get(
                'date_from', '.......'),
                _p.get_date().get('date_to', '.......')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Header Title 1
        row_title_body_pos = row_pos
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', 'Ngày tháng ghi sổ', None),
            ('col2', 2, 34, 'text', 'Chứng Từ', None),
            ('col3', 1, 45, 'text', 'Diễn giải', None),
            ('col4', 1, 12, 'text', 'Tài khoản ghi nợ', None),
            ('col5', 1, 12, 'text', 'Tài khoản ghi có', None),
            ('col6', 1, 22, 'text', 'Số tiền', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_bold_borderall, set_column_size=True)

        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 17, 'text', 'Số hiệu', None),
            ('col3', 1, 17, 'text', 'Ngày tháng', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 0,
                       0, 'Ngày tháng ghi sổ',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       3, 3, 'Diễn giải', self.normal_style_bold_borderall)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       4, 4, 'Tài khoản ghi nợ',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       5, 5, 'Tài khoản ghi có',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       6, 6, 'Số tiền', self.normal_style_bold_borderall)

        return row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        report_name = ''
        if _p.is_purchases_journal:
            report_name = u"SỔ NHẬT KÝ MUA HÀNG"
        else:
            report_name = u"SỔ NHẬT KÝ BÁN HÀNG"

        # call parent init utils.
        # set print sheet
        ws = super(account_purchases_journal_xls, self).generate_xls_report(
            _p, _xs, data, objects, wb, report_name)
        row_pos = self.generate_xls_header(
            _p, _xs, data, objects, wb, ws, 0, report_name)
        MAX_ROW = 65500
        count = 1

#         # account move lines
        get_lines_data = _p.get_lines_data()

        first_line_pos = row_pos
        sum_debit_amount = 0

        for line in get_lines_data:  # @UnusedVariable
            if row_pos > MAX_ROW:
                ws.flush_row_data()
                ws, row_pos = self.generate_worksheet(
                    _p, _xs, data, objects, wb, report_name, count)
                row_pos = self.generate_xls_header(
                    _p, _xs, data, objects, wb, ws, row_pos, report_name)
                count += 1
            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 450
            c_specs = map(lambda x: self.render(
                x, self.col_specs_template, 'lines'), self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data,
                row_style=self.normal_style_borderall, set_column_size=True)
            sum_debit_amount += line.get('amount', 0)
        last_line_pos = row_pos

        # Totals
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450

        if last_line_pos > first_line_pos:
            last_line_pos = last_line_pos - 1

        if not get_lines_data:
            # TODO: when there is no any records, we redefine cell style format
            self.col_specs_template['H'][
                'totals'] = self.col_specs_template['H']['totals_no_lines']
        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'totals'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.style_decimal)
#
        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(5)]
        c_specs = empty + [
            ('note1', 2, 0, 'text', 'Ngày .... tháng .... năm ....', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold']
        cell_footer_style = xlwt.easyxf(cell_format)

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', 'Người ghi sổ', None),
            ('col3', 1, 0, 'text', '', None),
            ('col6', 1, 16, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col9', 2, 0, 'text', 'Giám đốc', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['italic']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 1, 16, 'text', '(Ký, họ tên)', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)


account_purchases_journal_xls(
    'report.purchases_journal_report',
    'account.move',
    parser=account_purchases_journal_xls_parser
)
