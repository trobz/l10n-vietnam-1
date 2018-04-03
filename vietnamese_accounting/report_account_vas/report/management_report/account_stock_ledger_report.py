# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#
#    Copyright (c) 2014 Trobz (trobz.com). All rights reserved.
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program. If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

import xlwt
from datetime import datetime
from openerp.addons.report_xls.utils import rowcol_to_cell, _render # @UnresolvedImport
from openerp.addons.report_base_vn.report import report_base_vn
from .. import report_xls_utils
import logging
_logger = logging.getLogger(__name__)


class account_stock_ledger_xls_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(account_stock_ledger_xls_parser, self).__init__(cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_beginning_inventory': self.get_beginning_inventory,
            'get_account_info': self.get_account_info,
            'get_product_info': self.get_product_info,
            'is_purchases_journal': True,
            'get_wizard_data': self.get_wizard_data,
        })

    def get_company(self):

        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        address_list = [res.company_id.street or '',
                        res.company_id.street2 or '',
                        res.company_id.city or '',
                        res.company_id.state_id and res.company_id.state_id.name or '',
                        res.company_id.country_id and res.company_id.country_id.name or '',
        ]
        address_list = filter(None, address_list)
        address = ', '.join(address_list)
        vat = res.company_id.vat or ''
        return {'name': name, 'address': address, 'vat': vat }

    def get_wizard_data(self):

        result = {}
        datas = self.localcontext['data']
        if datas:
            result['fiscalyear'] = datas['form'] and datas['form']['fiscalyear_id'] or False
            result['target_move'] = datas['form'] and datas['form']['target_move'] or False
            result['account'] = datas['form'] and datas['form']['account'] and datas['form']['account'][0] or False
            result['location_id'] = datas['form'] and datas['form']['location_id'] and datas['form']['location_id'][0] or False
            result['product_id'] = datas['form'] and datas['form']['product_id'] and datas['form']['product_id'][0] or False
            result['location_name'] = datas['form'] and datas['form']['location_id'] and datas['form']['location_id'][1] or False
            result['filter'] = 'filter' in datas['form'] and datas['form']['filter'] or False
            if datas['form']['filter'] == 'filter_date':
                result['date_from'] = datas['form']['date_from']
                result['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                result['period_from'] = datas['form']['period_from']
                result['period_to'] = datas['form']['period_to']
        return result

    def get_account_info(self):

        account_id = self.get_wizard_data().get('account','') or False
        if account_id:
            res = self.pool.get('account.account').read(self.cr, self.uid, account_id, ['code','name'])
            return res
        return { 'code': '', 'name': ''}

    def get_product_info(self):
        product_id = self.get_wizard_data().get('product_id','') or False
        if product_id:
            product_obj = self.pool.get('product.product').browse(self.cr, self.uid, product_id)
            res = {
                   'default_code': product_obj.default_code or '',
                   'name': product_obj.name or '',
                   'product_uom': product_obj.uom_id and product_obj.uom_id.name or ''
                   }
            return res
        return { 'default_code': '', 'name': '', 'product_uom':''}

    def get_beginning_inventory(self):
        """
        Get the valuated total stock/cost of the lasted period
        """
        date_info = self.get_date()
        params  = { 'target_move': '', 'account_id': self.get_wizard_data()['account'],
                    'location_id': self.get_wizard_data()['location_id'],
                    'product_id': self.get_wizard_data()['product_id'],
                    'date_start': date_info['date_from_date'], 'date_end': date_info['date_to_date'] }
        SQL = """
            SELECT
                (SUM(CASE WHEN aml.debit > 0 THEN sm.product_qty ELSE 0 END) - SUM(CASE WHEN aml.debit > 0 THEN 0 ELSE sm.product_qty END)) as quantity_now,
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
                AND aml.product_id = %(product_id)s
                %(target_move)s
            GROUP BY aml.product_id
        """

        if self.get_wizard_data()['target_move'] == 'posted':
            params['target_move'] = " AND amv.state = 'posted'"


        _logger.info("SQL get the beginning inventory")
#         print SQL % params
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()

        return data

    def get_lines_data(self):

        """
        Get all account data
        debit_account is 11*
            credit_account
        """
        date_info = self.get_date()
        params  = { 'target_move': '', 'account_id': self.get_wizard_data()['account'],
                    'location_id': self.get_wizard_data()['location_id'],
                    'product_id': self.get_wizard_data()['product_id'],
                    'date_start': date_info['date_from_date'], 'date_end': date_info['date_to_date'] }


        SQL = """
        SELECT
            in_period.date_created,
            in_period.serial,
            in_period.effective_date,
            in_period.description,
            in_period.counterpart_account,
            in_period.price_unit,

            -- In Period
            COALESCE(in_period.quantity_in,0) as in_period_quantity_in,
            COALESCE(in_period.amount_in,0) as in_period_amount_in,

            COALESCE(in_period.quantity_out,0) as in_period_quantity_out,
            COALESCE(in_period.amount_out,0) as in_period_amount_out,

            -- End period
            (COALESCE(in_period.quantity_in,0) - COALESCE(in_period.quantity_out,0)) as end_period_quantity,
            (COALESCE(in_period.amount_in,0) - COALESCE(in_period.amount_out,0)) as end_period_amount

        FROM

            (SELECT
                aml.product_id,
                aml.date_created,
                aml.create_date,
                sp.name as serial,
                aml.date as effective_date,
                aml.name as description,
                acc.code as counterpart_account,
                -- sm.price_unit as price_unit,
                (CASE WHEN aml.debit > 0 THEN aml.debit/sm.product_qty ELSE aml.credit/sm.product_qty END) as price_unit,
                (CASE WHEN aml.debit > 0 THEN sm.product_qty ELSE 0 END) as quantity_in,
                (CASE WHEN aml.debit > 0 THEN 0 ELSE sm.product_qty END) as quantity_out,
                (aml.debit) as amount_in,
                (aml.credit) as amount_out,
                (aml.debit - aml.credit) as amount_now
            FROM
                account_move_line aml

            JOIN
                stock_move sm ON aml.stock_move_id = sm.id
            JOIN
                account_move amv ON amv.id = aml.move_id
            JOIN
                stock_picking sp ON sp.id = sm.picking_id

            LEFT JOIN
                account_move_line aml_cr ON aml.move_id = aml_cr.move_id

            LEFT JOIN
                account_account acc ON aml_cr.account_id = acc.id

            WHERE
                aml.account_id = %(account_id)s
                AND aml_cr.account_id != %(account_id)s
                AND (sm.location_id = %(location_id)s or sm.location_dest_id = %(location_id)s)
                AND aml.date >= '%(date_start)s' AND aml.date <= '%(date_end)s'
                AND aml.product_id = %(product_id)s
                %(target_move)s
            ) in_period

         ORDER BY create_date
        """

        if self.get_wizard_data()['target_move'] == 'posted':
            params['target_move'] = " AND amv.state = 'posted'"


#         print SQL % params
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data



class account_stock_ledger_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(account_stock_ledger_xls, self).__init__(name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A','B','C', 'D', 'E','F','G','H','I','K','L', 'M', 'N']
        self.col_specs_template = {
            'A': {
                'lines': [1, 0, 'date', _render("datetime.strptime(line.get('date_created',None)[:10],'%Y-%m-%d')"), None, self.style_date_right],
                'totals': [1, 0, 'text', None]},

            'B': {
                'lines': [1, 0, 'text', _render("line.get('serial')"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'C': {
                'lines': [1, 0, 'date', _render("datetime.strptime(line.get('effective_date',None)[:10],'%Y-%m-%d')"), None, self.style_date_right],
                'totals': [1, 0, 'text', None]},

            'D': {
                'lines': [1, 0, 'text', _render("line.get('description')"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'E': {
                'lines': [1, 0, 'text', _render("line.get('counterpart_account')"), None, self.normal_style_right_borderall],
                'totals': [1, 0, 'text', None]},

            'F': {
                'lines': [1, 0, 'number', _render("line.get('price_unit',None)"), None, self.style_decimal],
                'totals': [1, 0, 'text', None]},

            'G': {
                'lines': [1, 0, 'number', _render("line.get('in_period_quantity_in',None)"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'H': {
                'lines': [1, 0, 'number', _render("line.get('in_period_amount_in',None)"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'I': {
                'lines': [1, 0, 'number', _render("line.get('in_period_quantity_out',None)"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'K': {
                'lines': [1, 0, 'number', _render("line.get('in_period_amount_out')"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'L': {
                'lines': [1, 0, 'number', _render("line.get('end_period_quantity')"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'M': {
                'lines': [1, 0, 'number', _render("line.get('end_period_amount')"), None, self.style_decimal],
                'totals': [1, 0, 'number', None, None, self.style_decimal_bold]},

            'N': {
                'lines': [1, 0, 'text', '', None, None],
                'totals': [1, 0, 'text', '', None, None]},
        }


    def generate_xls_report(self, _p, _xs, data, objects, wb):
        report_name = 'SỔ CHI TIẾT VẬT LIỆU, DỤNG CỤ (SẢN PHẨM, HÀNG HÓA)'

        # call parent init utils.
        # set print sheet
        ws = super(account_stock_ledger_xls, self).generate_xls_report(_p, _xs, data, objects, wb, report_name)

        row_pos = 0

        cell_address_style = self.get_cell_style(['bold', 'wrap', 'left'])
        # Title address 1
        c_specs = [
            ('company_name', 6, 0, 'text', u'Đơn vị: %s' % _p.get_company()['name']  , '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 2
        c_specs = [
            ('company_name', 6, 0, 'text', u'Địa chỉ: %s' % _p.get_company()['address'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('company_name', 6, 0, 'text', u'MST: %s' % _p.get_company()['vat'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

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
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "SỔ NHẬT KÝ MUA HÀNG"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_title_style = self.get_cell_style(['bold', 'wrap', 'center', 'middle', 'fontsize_350'])

        c_specs = [
            ('payment_journal', 13, 0, 'text', report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Loại tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 13, 0, 'text', u'Loại tài khoản: %s' % (_p.get_account_info().get('code','')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 13, 0, 'text', u'Tên tài khoản: %s' % (_p.get_account_info().get('name','')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên kho"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 13, 0, 'text', u'Tên kho: %s' % (_p.get_wizard_data().get('location_name','')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Mã nguyên liệu, vật liệu, công cụ, dụng cu (sản phẩm, hàng hóa)"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 13, 0, 'text', u'Mã nguyên liệu, vật liệu, công cụ, dụng cụ (sản phẩm, hàng hóa): %s' % (_p.get_product_info().get('default_code','') or ''))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên nguyên liệu, vật liệu, công cụ, dụng cu (sản phẩm, hàng hóa)"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('title', 13, 0, 'text', u'Tên nguyên liệu, vật liệu, công cụ, dụng cụ (sản phẩm, hàng hóa): %s' % (_p.get_product_info().get('name','') or ''))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('from_to', 13, 0, 'text', u'Từ %s đến %s' % (_p.get_date().get('date_from','.......'),_p.get_date().get('date_to','.......')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(11)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)


        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(11)]
        c_specs = c_specs + [('measure_unit',2, 12, 'text', u'Đơn vị tính: %s' % (_p.get_product_info().get('product_uom','') or ''), None, self.normal_style_right)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)


        # Header Title 1
        row_title_body_pos = row_pos
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450
        c_specs = [
            ('col1', 1, 10, 'text', 'Ngày tháng ghi sổ', None),
            ('col2', 2, 24, 'text', 'Chứng từ', None ),
            ('col3', 1, 24, 'text', 'Diễn giải', None),
            ('col4', 1, 12, 'text', 'TK đối ứng', None),
            ('col5', 1, 12, 'text', 'Đơn giá', None),
            ('col6', 2, 22, 'text', 'Nhập', None),
            ('col7', 2, 22, 'text', 'Xuất', None),
            ('col8', 2, 22, 'text', 'Tồn', None),
            ('col9', 1, 22, 'text', 'Ghi chú', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)


        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', 'Số hiệu', None),
            ('col3', 1, 12, 'text', 'Ngày tháng', None),
            ('col4', 1, 24, 'text', '', None),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 12, 'text', 'Số lượng', None),
            ('col8', 1, 12, 'text', 'Thành tiền', None),
            ('col9', 1, 12, 'text', 'Số lượng', None),
            ('col10', 1, 12, 'text', 'Thành tiền', None),
            ('col11', 1, 12, 'text', 'Số lượng', None),
            ('col12', 1, 12, 'text', 'Thành tiền', None),

        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 0, 0, 'Ngày tháng ghi sổ', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 3, 3, 'Diễn giải', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 4, 4, 'TK đối ứng', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 5, 5, 'Đơn giá', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 12, 12, 'Ghi chú', self.normal_style_bold_borderall )
        ws.write_merge(0, 2, 6, 12, '''Mẫu số S10 – DN
(Ban hành theo QĐ số 15/2006/QĐ-BTC, Ngày 20/03/2006 của Bộ trưởng BTC)''', self.normal_style )



        get_beginning_inventory = _p.get_beginning_inventory()
        if not get_beginning_inventory:
            get_beginning_inventory = {}
        else:
            get_beginning_inventory = get_beginning_inventory[0]

        # The beginning inventory data
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', 'Số dư đầu kỳ', None, self.normal_style_bold),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 12, 'text', '', None),
            ('col8', 1, 12, 'text', '', None),
            ('col9', 1, 12, 'text', '', None),
            ('col10', 1, 12, 'text', '', None),
            ('col11', 1, 12, 'number', get_beginning_inventory.get('quantity_now',0), None, self.style_decimal),
            ('col12', 1, 12, 'number', get_beginning_inventory.get('amount_now',0), None, self.style_decimal),
            ('col13', 1, 12, 'text', '', None),

        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_borderall)

        # account move lines
        get_lines_data = _p.get_lines_data()

        first_line_pos = row_pos
        previous_end_amount = get_beginning_inventory.get('amount_now',0)
        previous_end_quantity = get_beginning_inventory.get('quantity_now',0)
        for line in get_lines_data: # @UnusedVariable
            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 450
            # caculate ending stock/amount for current stock_move
            previous_end_amount = previous_end_amount + line.get('end_period_amount',0)
            previous_end_quantity = previous_end_quantity + line.get('end_period_quantity',0)
            line.update({ 'end_period_quantity': previous_end_quantity, 'end_period_amount': previous_end_amount })
            c_specs = map(lambda x: self.render(x, self.col_specs_template, 'lines'), self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_borderall)
        last_line_pos = row_pos

        # Totals
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450

        if last_line_pos > first_line_pos:
            last_line_pos = last_line_pos - 1

        # sum for thes columns

        sum_columns = ['G','H','I','K']
        self.col_specs_template['D']['totals'] = [1, 0, 'text', 'Tổng Cộng', None, self.normal_style_bold_borderall]
        if get_lines_data:
            _logger.info(" >>>>>>>>>>>>>>>>>>> Start printing totals row <<<<<<<<<<<<<<<<<<<<<<")
            # if there is some line
            for column in sum_columns:
                value_start = rowcol_to_cell(first_line_pos, self.wanted_list.index(column))
                value_stop = rowcol_to_cell(last_line_pos, self.wanted_list.index(column))
                self.col_specs_template[column]['totals'] = [1, 0, 'number', None, 'SUM(%s:%s)' % (value_start, value_stop),  self.style_decimal_bold]
        else:
            # TODO: we use this case because we face weird error for the sum row
            # it seems cache somewhere ??
            _logger.info(" >>>>>>>>>>>>>>>>>>> Start reset totals row <<<<<<<<<<<<<<<<<<<<<<")
            # set null for these columns
            for column in sum_columns:
            # TODO: when there is no any records, we redefine cell style format
                self.col_specs_template[column]['totals'] = [1, 0, 'number', None, None, self.style_decimal_bold]

        self.col_specs_template['L']['totals'] = [1, 0, 'number', previous_end_quantity, None, self.style_decimal_bold]
        self.col_specs_template['M']['totals'] = [1, 0, 'number', previous_end_amount, None, self.style_decimal_bold]

        c_specs = map(lambda x: self.render(x, self.col_specs_template, 'totals'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.style_decimal)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = empty + [
            ('note1', 2, 0, 'text','Ngày .... tháng .... năm ....', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold']
        cell_footer_style = xlwt.easyxf(cell_format)

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', 'Người ghi sổ', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 3, 16, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 1, 0, 'text', '', None),
            ('col10', 2, 0, 'text', 'Giám đốc', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['italic']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 3, 16, 'text', '(Ký, họ tên)', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 1, 0, 'text', '', None),
            ('col10', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=cell_footer_style)

account_stock_ledger_xls('report.stock_ledger_report','stock.move', parser=account_stock_ledger_xls_parser)

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
