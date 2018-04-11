# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import xlwt
from datetime import datetime
from openerp.addons.report_xls.utils import rowcol_to_cell, _render
from openerp.addons.report_base_vn.report import report_base_vn
from .. import report_xls_utils
import logging
_logger = logging.getLogger(__name__)


class account_stock_balance_xls_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(account_stock_balance_xls_parser, self).__init__(
            cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_account_info': self.get_account_info,
            'is_purchases_journal': True,
            'get_wizard_data': self.get_wizard_data,
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
            result['location_id'] = datas['form'] and datas['form'][
                'location_id'] and datas['form']['location_id'][0] or False
            result['location_name'] = datas['form'] and datas['form'][
                'location_id'] and datas['form']['location_id'][1] or False
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
                  'account_id': self.get_wizard_data()['account'],
                  'location_id': self.get_wizard_data()['location_id'],
                  'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date']}

        SQL = """
        /*
        When receiving product,
            each stock_move will have only one Journal Entry
        In 2 Journal Items of this Journal,
            we have stock_move_id is stock_move this item belong to.
        Depend on stock_move and account_move_line,
            we will get all info for stock_balance report.
        */

        SELECT

            COALESCE(pp1.product_code,pp2.product_code) AS product_code,
            COALESCE(pp1.product_name,pp2.product_name) AS product_name,

            -- Beginning Period
            (COALESCE(start_period.quantity_in,0) - COALESCE(start_period.quantity_out,0)) as start_period_quantity,
            (COALESCE(start_period.amount_in,0) - COALESCE(start_period.amount_out,0)) as start_period_amount,

            -- In Period
            COALESCE(in_period.quantity_in,0) as in_period_quantity_in,
            COALESCE(in_period.amount_in,0) as in_period_amount_in,

            COALESCE(in_period.quantity_out,0) as in_period_quantity_out,
            COALESCE(in_period.amount_out,0) as in_period_amount_out,

            -- End period
            (COALESCE(start_period.quantity_in,0) - COALESCE(start_period.quantity_out,0) + COALESCE(in_period.quantity_in,0) - COALESCE(in_period.quantity_out,0)) as end_period_quantity,
            (COALESCE(start_period.amount_in,0) - COALESCE(start_period.amount_out,0) + COALESCE(in_period.amount_in,0) - COALESCE(in_period.amount_out,0)) as end_period_amount

        FROM
        (

            (SELECT
                  aml.product_id,
                  SUM(CASE WHEN aml.debit > 0 THEN sm.product_qty ELSE 0 END) as quantity_in,
                  SUM(CASE WHEN aml.debit > 0 THEN 0 ELSE sm.product_qty END) as quantity_out,
                  SUM(aml.debit) as amount_in, sum(aml.credit) as amount_out,
                  SUM(aml.debit - aml.credit) as amount_now
            FROM
                account_move_line aml
            JOIN
                stock_move sm ON aml.stock_move_id = sm.id
            JOIN
                account_move amv ON amv.id= aml.move_id
            WHERE
                aml.account_id = %(account_id)s
                AND (sm.location_id = %(location_id)s or sm.location_dest_id = %(location_id)s)
                AND aml.date < '%(date_start)s'
                %(target_move)s
            GROUP BY aml.product_id ) start_period

        FULL JOIN  -- JOIN to get IN_PERIOD Table

            (SELECT
                aml.product_id,
                SUM(CASE WHEN aml.debit > 0 THEN sm.product_qty ELSE 0 END) as quantity_in,
                SUM(CASE WHEN aml.debit > 0 THEN 0 ELSE sm.product_qty END) as quantity_out,
                SUM(aml.debit) as amount_in,
                SUM(aml.credit) as amount_out,
                SUM(aml.debit - aml.credit) as amount_now
            FROM
                account_move_line aml
            JOIN
                stock_move sm ON aml.stock_move_id = sm.id
            JOIN
                account_move amv ON amv.id= aml.move_id
            WHERE
                aml.account_id = %(account_id)s
                AND (sm.location_id = %(location_id)s or sm.location_dest_id = %(location_id)s)
                AND aml.date >= '%(date_start)s' AND aml.date <= '%(date_end)s'
                %(target_move)s
            GROUP BY
                aml.product_id) in_period ON in_period.product_id=start_period.product_id

        LEFT JOIN  -- JOIN with product_product table to get product info
            (   SELECT id,
                default_code as product_code,
                name_template as product_name
            FROM product_product ) pp1 ON pp1.id=start_period.product_id
        -- LEFT JOIN with product_product table again in case there is no recored in start_period table

        LEFT JOIN (SELECT id,
                default_code as product_code,
                name_template as product_name
            FROM product_product ) pp2 ON pp2.id=in_period.product_id
        )

         ORDER BY pp1.id, pp2.id

        """

        if self.get_wizard_data()['target_move'] == 'posted':
            params['target_move'] = " AND amv.state = 'posted'"

        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data


class account_stock_balance_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(account_stock_balance_xls, self).__init__(
            name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E', 'G', 'H', 'I', 'K', 'L']
        self.col_specs_template = {
            'A': {
                'lines': [1, 0, 'text', _render("line.get('product_code')"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'B': {
                'lines': [1, 0, 'text', _render("line.get('product_name')"), None, self.normal_style_left_borderall]},

            'C': {
                'lines': [1, 0, 'number', _render("line.get('start_period_quantity')"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'D': {
                'lines': [1, 0, 'number', _render("line.get('start_period_amount')"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'E': {
                'lines': [1, 0, 'number', _render("line.get('in_period_quantity_in',None)"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'G': {
                'lines': [1, 0, 'number', _render("line.get('in_period_amount_in',None)"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'H': {
                'lines': [1, 0, 'number', _render("line.get('in_period_quantity_out',None)"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'I': {
                'lines': [1, 0, 'number', _render("line.get('in_period_amount_out',None)"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'K': {
                'lines': [1, 0, 'number', _render("line.get('end_period_quantity')"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'L': {
                'lines': [1, 0, 'number', _render("line.get('end_period_amount')"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},
        }

    def generate_xls_report(self, _p, _xs, data, objects, wb):
        report_name = 'BẢNG TỔNG HỢP CHI TIẾT VẬT LIỆU, DỤNG CỤ (SẢN PHẨM, HÀNG HÓA)'

        # call parent init utils.
        # set print sheet
        ws = super(account_stock_balance_xls, self).generate_xls_report(
            _p, _xs, data, objects, wb, report_name)
        row_pos = 0

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
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 1, 0, 'text', '', None),
            ('col7', 1, 0, 'text', '', None),
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
            ('payment_journal', 10, 0, 'text', report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Loại tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 10, 0, 'text', u'Loại tài khoản: %s' %
             (_p.get_account_info().get('code', '')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 10, 0, 'text', u'Tên tài khoản: %s' %
             (_p.get_account_info().get('name', '')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên kho"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 10, 0, 'text', u'Tên kho: %s' %
             (_p.get_wizard_data().get('location_name', '')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('from_to', 10, 0, 'text', u'Từ %s đến %s' % (_p.get_date().get(
                'date_from', '.......'), _p.get_date().get('date_to', '.......')))
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
        ws.row(row_pos).height = 450
        c_specs = [
            ('col1', 1, 12, 'text',
             'Mã nguyên liệu, vật liệu, công cụ, dụng cụ (sản phẩm, hàng hóa)', None),
            ('col2', 1, 34, 'text',
             'Tên nguyên liệu, vật liệu, công cụ, dụng cụ (sản phẩm, hàng hóa)', None),
            ('col3', 2, 45, 'text', 'Tồn đầu kỳ', None),
            ('col4', 2, 12, 'text', 'Nhập trong kỳ', None),
            ('col5', 2, 12, 'text', 'Xuất trong kỳ', None),
            ('col6', 2, 22, 'text', 'Tồn cuối kỳ', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)

        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450
        c_specs = [
            ('col1', 1, 22, 'text', '', None),
            ('col2', 1, 22, 'text', '', None),
            ('col3', 1, 12, 'text', 'Số lượng', None),
            ('col4', 1, 12, 'text', 'Thành tiền', None),
            ('col5', 1, 12, 'text', 'Số lượng', None),
            ('col6', 1, 12, 'text', 'Thành tiền', None),
            ('col7', 1, 12, 'text', 'Số lượng', None),
            ('col8', 1, 12, 'text', 'Thành tiền', None),
            ('col9', 1, 12, 'text', 'Số lượng', None),
            ('col10', 1, 12, 'text', 'Thành tiền', None),

        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 0, 0,
                       'Mã nguyên liệu, vật liệu, công cụ, dụng cụ (sản phẩm, hàng hóa)', self.normal_style_bold_borderall)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 1, 1,
                       'Tên nguyên liệu, vật liệu, công cụ, dụng cu (sản phẩm, hàng hóa)', self.normal_style_bold_borderall)


#         # account move lines
        get_lines_data = _p.get_lines_data()

        first_line_pos = row_pos
        for line in get_lines_data:  # @UnusedVariable
            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 450
            c_specs = map(lambda x: self.render(
                x, self.col_specs_template, 'lines'), self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data, row_style=self.normal_style_borderall)
        last_line_pos = row_pos

        # Totals
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450

        if last_line_pos > first_line_pos:
            last_line_pos = last_line_pos - 1

        # sum for thes columns

        sum_columns = ['A', 'C', 'D', 'E', 'G', 'H', 'I', 'K', 'L']
        self.col_specs_template['A']['totals'] = [
            2, 0, 'text', 'Tổng Cộng', None, self.normal_style_bold_borderall]
        if get_lines_data:
            # if there is some line
            for column in sum_columns[1:]:
                value_start = rowcol_to_cell(
                    first_line_pos, self.wanted_list.index(column))
                value_stop = rowcol_to_cell(
                    last_line_pos, self.wanted_list.index(column))
                self.col_specs_template[column]['totals'] = [1, 0, 'number', None, 'SUM(%s:%s)' % (
                    value_start, value_stop),  self.style_decimal_bold]
        else:
            # set null for these columns
            for column in sum_columns[1:]:
                # TODO: when there is no any records, we redefine cell style
                # format
                self.col_specs_template[column]['totals'] = [
                    1, 0, 'number', None, None, self.style_decimal_bold]
        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'totals'), sum_columns)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.style_decimal)

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
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(8)]
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
            ('col6', 3, 16, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
            ('col10', 2, 0, 'text', 'Giám đốc', None),
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
            ('col6', 3, 16, 'text', '(Ký, họ tên)', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
            ('col10', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)


account_stock_balance_xls('report.stock_balance_report',
                          'stock.move', parser=account_stock_balance_xls_parser)
