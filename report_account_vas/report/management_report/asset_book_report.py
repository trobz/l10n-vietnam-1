# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import xlwt  # @UnresolvedImport
from datetime import datetime
from openerp.addons.report_xls.report_xls import report_xls
from openerp.addons.report_xls.utils import rowcol_to_cell, _render
from openerp.tools.safe_eval import safe_eval
from openerp.tools.translate import translate, _
from openerp.addons.report_base_vn.report import report_base_vn
import logging
_logger = logging.getLogger(__name__)

_ir_translation_name = 'move.line.list.xls'


class asset_book_xls_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(asset_book_xls_parser, self).__init__(
            cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'asset_data': self.get_asset_data,
            'get_asset_category': self.get_asset_category,
            'get_report_param': self.get_report_param
        })

    def get_wizard_data(self):
        result = {}
        datas = self.localcontext['data']
        if datas:
            result['fiscalyear'] = datas['form'] and datas[
                'form']['fiscalyear_id'] or False
            result['chart_account_id'] = datas['form'] and datas[
                'form']['chart_account_id'] or False
            result['asset_category_id'] = datas['form'] and datas[
                'form']['asset_category_id'] and [datas[
                    'form']['asset_category_id'][0]] or \
                self.pool.get('account.asset.category').search(
                    self.cr, self.uid, [])
            result['filter'] = 'filter' in datas[
                'form'] and datas['form']['filter'] or False
            if datas['form']['filter'] == 'filter_date':
                result['date_from'] = datas['form']['date_from']
                result['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                result['period_from'] = datas['form']['period_from']
                result['period_to'] = datas['form']['period_to']
        return result

    def get_asset_category(self):
        return self.localcontext['data'] and self.localcontext['data']['form'] \
            and self.localcontext['data']['form']['asset_category_id'] and self.localcontext['data']['form']['asset_category_id'][1] or '............'

    def get_asset_data(self):
        """
        Get all asset data
        """
        date_info = self.get_date()
        category_ids = self.get_wizard_data()['asset_category_id']
        params = {'category_ids': tuple(
            category_ids + [-1, -1]),
            'date_start': date_info['date_from_date'],
            'date_end': date_info['date_to_date']}

        SQL = """
            SELECT COALESCE(ai.number,asset.code) as number,
            COALESCE(ai.creation_date,purchase_date) as date,
            asset.reference as reference,basic_specification,
            COALESCE(country.name,'') as manufacturer,
            manufacture_year,department,start_date_in_use,
            capacity,date_of_liquidation,reason,
            aml.date_created AS effective_date,move_check,
            purchase_value, adl.remaining_value,
            (asset.purchase_value - adl.remaining_value) as amount_depreciated,
            CASE WHEN method = 'linear'
                THEN 0
                ELSE method_progress_factor
            END AS deprecition_rate,

            CASE WHEN method = 'linear'
                THEN (SELECT amount
                        FROM account_asset_depreciation_line
                        WHERE asset.id=asset_id LIMIT 1)
                ELSE 0
            END AS current_deprecition

            FROM account_asset_asset asset
            LEFT JOIN account_invoice as ai ON ai.number=asset.code
            LEFT JOIN (SELECT min(date_created) as date_created,asset_id
                        FROM account_move_line GROUP BY asset_id) AS aml
                ON aml.asset_id=asset.id
            LEFT JOIN (SELECT MIN(remaining_value)
                        as remaining_value ,asset_id,move_check
                FROM account_asset_depreciation_line
                GROUP BY asset_id,move_check) AS adl
                ON adl.asset_id=asset.id AND move_check

            LEFT JOIN res_country country ON country.id=asset.manufacturer
            WHERE category_id in %(category_ids)s
            AND ai.creation_date >= '%(date_start)s'
            AND ai.creation_date <= '%(date_end)s'
            AND asset.state != 'draft'
        """
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data

    def get_report_param(self, param=''):
        report_param = self.pool.get('ir.config_parameter').get_param(
            self.cr, self.uid, 'account_asset_book_report_parameter')
        report_param = safe_eval(report_param)
        return report_param.get(param, '')

    def _(self, src):
        lang = self.context.get('lang', 'en_US')
        return translate(
            self.cr, _ir_translation_name, 'report', lang, src) or src


class asset_book_xls(report_xls, report_base_vn.Parser):

    def __init__(
            self, name, table, rml=False,
            parser=False, header=True, store=False):
        super(asset_book_xls, self).__init__(
            name, table, rml, parser, header, store)

        # Cell Styles
        date_format = 'DD-MM-YYYY'
        decimal_format = '#,##0'
        _xs = self.xls_styles
        # header
        rh_cell_format = _xs['bold'] + _xs['fill'] + \
            _xs['borders_all'] + _xs['wrap'] + _xs['middle']
        self.rh_cell_style = xlwt.easyxf(rh_cell_format)
        self.rh_cell_style_center = xlwt.easyxf(rh_cell_format + _xs['center'])
        self.rh_cell_style_right = xlwt.easyxf(rh_cell_format + _xs['right'])
        # lines
        aml_cell_format = _xs['borders_all'] + _xs['wrap'] + _xs['middle']
        self.aml_cell_style = xlwt.easyxf(aml_cell_format)
        self.aml_cell_style_center = xlwt.easyxf(
            aml_cell_format + _xs['center'])
        self.aml_cell_style_date = xlwt.easyxf(
            aml_cell_format + _xs['left'] + _xs['center'], num_format_str=date_format)
        self.aml_cell_style_decimal = xlwt.easyxf(
            aml_cell_format + _xs['right'], num_format_str=decimal_format)
        # totals
        rt_cell_format = _xs['bold'] + _xs['fill'] + \
            _xs['borders_all'] + _xs['wrap'] + _xs['middle']
        self.rt_cell_style = xlwt.easyxf(rt_cell_format)
        self.rt_cell_style_right = xlwt.easyxf(rt_cell_format + _xs['right'])
        self.rt_cell_style_decimal = xlwt.easyxf(
            rt_cell_format + _xs['right'], num_format_str=decimal_format)

        self.asset_xls_styles = {
            'normal': '',
            'bold': '',
            'underline': 'font: underline true;',
        }

        # normal
        self.cell_style_normal = xlwt.easyxf(
            self.asset_xls_styles['normal'] + _xs['borders_all'] + _xs['wrap'] + _xs['center'] + _xs['middle'])
        self.cell_style_normal_borderless = xlwt.easyxf(
            self.asset_xls_styles['normal'] + _xs['wrap'] + _xs['center'] + _xs['middle'])
        self.cell_style_normal_borderless_italic = xlwt.easyxf(
            self.asset_xls_styles['normal'] + _xs['wrap'] + _xs['center'] + _xs['middle'] + _xs['italic'])
        cell_total_style = xlwt.easyxf(
            _xs['wrap'] + _xs['center'] + _xs['bold'] + _xs['borders_all'])

        # center
        self.cell_style_center = xlwt.easyxf(
            _xs['center'] + _xs['borders_all'] + _xs['wrap'] + _xs['middle'])

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E',
                            'G', 'H', '1', '2', '3', '4', 'I', 'K', 'L']
        self.col_specs_template = {
            'A': {
                'header': [1, 8, 'text', _render("('A')")],
                'lines': [1, 0, 'number', _render("next_number()")],
                'totals': [1, 0, 'text', None]},
            'B': {
                'header': [1, 16, 'text', _render("('B')")],
                'lines': [1, 0, 'text', _render("line.get('number','') or ''")],
                'totals': [1, 0, 'text', None]},
            'C': {
                'header': [1, 10, 'text', _render("('C')")],
                'lines': [1, 0, 'date', _render("datetime.strptime(line.get('date',None)[:10],'%Y-%m-%d')"), None, self.aml_cell_style_date],
                'totals': [1, 0, 'text', None]},
            'D': {
                'header': [1, 13, 'text', _render("('D')")],
                'lines': [1, 0, 'text', _render("line.get('basic_specification','') or ''")],
                'totals': [1, 0, 'text', 'Cộng', None, cell_total_style]},
            'E': {
                'header': [1, 12, 'text', _render("('E')")],
                'lines': [1, 0, 'text', _render("line.get('manufacturer',None)")],
                'totals': [1, 0, 'text', None]},
            'G': {
                'header': [1, 10, 'text', _render("('G')")],
                'lines': [1, 0, _render("line.get('effective_date','') and 'date' or 'text'"), _render("line.get('effective_date','') and datetime.strptime(line.get('effective_date',''),'%Y-%m-%d') or None"), None, self.aml_cell_style_date],
                'totals': [1, 0, 'text', None]},
            'H': {
                'header': [1, 8, 'text', _render("('H')")],
                'lines': [1, 0, 'text', _render("line.get('reference','') or ''"), None],
                'totals': [1, 0, 'text', None]},
            '1': {
                'header': [1, 15, 'text', _render("('1')")],
                'lines': [1, 0, 'number', _render("line.get('purchase_value',0)"), None, self.aml_cell_style_decimal],
                'totals': [1, 0, 'number', None, _render("sum_purchase_value_formula"), None, self.aml_cell_style_decimal]},
            '2': {
                'header': [1, 10, 'text', _render("('2')")],
                'lines': [1, 0, 'text', _render("line.get('deprecition_rate','') == 0 and '-' or str(line.get('deprecition_rate')*100)"), None],
                'totals': [1, 0, 'text', None]},
            '3': {
                'header': [1, 10, 'text', _render("('3')"), None],
                'lines': [1, 0, 'text', _render("line.get('current_deprecition','') == 0 and '-' or str(line.get('current_deprecition'))"), None],
                'totals': [1, 0, 'text', None, None]},
            '4': {
                'header': [1, 13, 'text', _render("('4')"), None],
                'lines': [1, 0, 'number', _render("line.get('amount_depreciated',0)"), None, self.aml_cell_style_decimal],
                'totals': [1, 0, 'text', None, None]},
            'I': {
                'header': [1, 8, 'text', _render("('I')"), None],
                'lines': [1, 0, 'text', None, None, self.aml_cell_style_decimal],
                'totals': [1, 0, 'text', 'x', None]},
            'K': {
                'header': [1, 11, 'text', _render("('K')"), None],
                'lines': [1, 0, 'text', None, None, self.aml_cell_style_center],
                'totals': [1, 0, 'text', 'x']},
            'L': {
                'header': [1, 11, 'text', _render("('L')"), None],
                'lines': [1, 0, 'text', None, None, self.aml_cell_style_center],
                'totals': [1, 0, 'text', 'x']},
        }

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        # report_name = objects[0]._description or objects[0]._name
        report_name = _("Sổ tài sản cố định")

        ws = wb.add_sheet(report_name[:31], cell_overwrite_ok=True)
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_width_to_pages = 1
        ws.fit_num_pages = 1
        ws.show_grid = 0
        row_pos = 0

        # set print header/footer
        ws.header_str = self.xls_headers['standard']
        ws.footer_str = self.xls_footers['standard']
        cell_address_format = _xs['bold'] + _xs['wrap'] + _xs['left']
        cell_address_style = xlwt.easyxf(cell_address_format)
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold']
        cell_footer_style = xlwt.easyxf(cell_format)

        # Title 1
        c_specs = [
            ('company_name', 3, 0, 'text', u'Đơn vị: %s' %
             _p.get_company()['name'], '', cell_address_style),
            ('empty1', 1, 0, 'text', ''),
            ('empty2', 1, 0, 'text', ''),
            ('empty3', 1, 0, 'text', ''),
            ('empty4', 1, 0, 'text', ''),
            ('empty5', 1, 0, 'text', ''),
            ('empty6', 1, 0, 'text', ''),
            ('empty7', 1, 0, 'text', ''),
            ('form_serial', 4, 0, 'text', _p.get_report_param(
                'format_form'), '', cell_footer_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        # Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('company_name', 3, 0, 'text', u'Địa chỉ: %s' %
             _p.get_company()['address'], '', cell_address_style),
            ('empty1', 1, 0, 'text', ''),
            ('empty2', 1, 0, 'text', ''),
            ('empty3', 1, 0, 'text', ''),
            ('empty4', 1, 0, 'text', ''),
            ('empty5', 1, 0, 'text', ''),
            ('empty6', 1, 0, 'text', ''),
            ('empty7', 1, 0, 'text', ''),
            ('form_serial', 4, 0, 'text', _p.get_report_param(
                'form_formated_by_rule'), '',
             self.cell_style_normal_borderless)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        # Title "sổ tài sản cố định"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_format = _xs['bold'] + _xs['wrap'] + \
            _xs['center'] + 'font: height 370;'
        cell_title_style = xlwt.easyxf(cell_format)

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(5)]
        c_specs = empty + [
            ('asset_book', 4, 10, 'text', 'SỔ TÀI SẢN CỐ ĐỊNH',
                None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = empty + [
            ('asset_book', 4, 0, 'text', 'Từ %s Đến %s' % (_p.get_date().get(
                'date_from', '.......'),
                _p.get_date().get('date_to', '.......')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.cell_style_normal_borderless_italic)

        # Title "loại tài sản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['bold'] + _xs['wrap'] + _xs['center']
        cell_title_style = xlwt.easyxf(cell_format)
        c_specs = empty + [
            ('from_to', 4, 10, 'text', u'Loại tài sản: %s' %
             _p.get_asset_category(), None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_title_style)

        # Add 2 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        # title 3
        c_specs = [
            ('number', 1, 0, 'text', 'STT', None, self.cell_style_normal),
            ('increase_asset', 7, 0, 'text',
             'Ghi tăng TSCĐ', None, self.cell_style_normal),
            ('depreciation_asset', 3, 0, 'text',
             'Khấu hao TSCĐ', None, self.cell_style_normal),
            ('decrease_asset', 3, 0, 'text',
             'Ghi giảm TSCĐ', None, self.cell_style_normal)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal)

        # Title 4
        row_title_body_pos = row_pos
        c_specs = [
            ('col1', 1, 0, 'text', '', None, self.cell_style_center),
            ('col2', 2, 0, 'text', 'Chứng Từ', None, self.cell_style_center),
            ('col3', 1, 0, 'text', 'Tên, đặc điểm, ký hiệu TSCĐ',
             None, self.cell_style_center),
            ('col4', 1, 0, 'text', 'Nước sản xuất',
                None, self.cell_style_center),
            ('col5', 1, 0, 'text', 'Tháng năm đưa vào sử dụng',
             None, self.cell_style_center),
            ('col6', 1, 0, 'text', 'Số hiệu TSCĐ',
                None, self.cell_style_center),
            ('col7', 1, 0, 'text', 'Nguyên giá TSCĐ',
                None, self.cell_style_center),
            ('col8', 2, 0, 'text', 'Khấu hao',
                None, self.cell_style_center),
            ('col9', 1, 0, 'text', 'Khấu hao đã tính đến khi giảm TSCĐ',
             None, self.cell_style_center),
            ('col10', 2, 0, 'text', 'Chứng từ', None, self.cell_style_center),
            ('col11', 1, 0, 'text', 'Lý do giảm TSCĐ',
                None, self.cell_style_center)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal)

        # Title 5
        c_specs = [
            ('col1', 1, 0, 'text', '', None, self.cell_style_center),
            ('col2', 1, 0, 'text', 'Số hiệu', None, self.cell_style_center),
            ('col3', 1, 0, 'text', 'Ngày tháng', None, self.cell_style_center),
            ('col4', 1, 0, 'text', '', None, self.cell_style_center),
            ('col5', 1, 0, 'text', '', None, self.cell_style_center),
            ('col6', 1, 0, 'text', '', None, self.cell_style_center),
            ('col7', 1, 0, 'text', '', None, self.cell_style_center),
            ('col8', 1, 0, 'text', '', None, self.cell_style_center),
            ('col9', 1, 0, 'text', 'Tỷ lệ (%) khấu hao',
             None, self.cell_style_center),
            ('col10', 1, 0, 'text', 'Mức khấu hao',
                None, self.cell_style_center),
            ('col11', 1, 0, 'text', '', None, self.cell_style_center),
            ('col12', 1, 0, 'text', 'Số hiệu', None, self.cell_style_center),
            ('col13', 1, 0, 'text', 'Ngày, tháng, năm',
                None, self.cell_style_center),
            ('col14', 1, 0, 'text', '', None, self.cell_style_center)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal)

        # merge cell
        ws.write_merge(row_title_body_pos - 1, row_title_body_pos +
                       1, 0, 0, 'STT', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 3,
                       3, 'Tên, đặc điểm, ký hiệu TSCĐ',
                       self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       4, 4, 'Nước sản xuất', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 5,
                       5, 'Tháng năm đưa vào sử dụng',
                       self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       6, 6, 'Số hiệu TSCĐ', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       7, 7, 'Nguyên giá TSCĐ', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1, 10, 10,
                       'Khấu hao đã tính đến khi giảm TSCĐ',
                       self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos + 1,
                       13, 13, 'Lý do giảm TSCĐ', self.aml_cell_style_center)

        ws.row(row_title_body_pos + 1).height_mismatch = True
        ws.row(row_title_body_pos + 1).height = 256 * 2
        # Column headers
        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'header'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.aml_cell_style_center, set_column_size=True)
        ws.set_horz_split_pos(row_pos)

        # account move lines
        asset_data = _p.asset_data()
        first_line_pos = row_pos
        for line in asset_data:  # @UnusedVariable
            c_specs = map(lambda x: self.render(
                x, self.col_specs_template, 'lines'), self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data, row_style=self.aml_cell_style_center)
        last_line_pos = row_pos
        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(14)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal)

        # Totals
        purchase_value_start = rowcol_to_cell(first_line_pos, 7)
        purchase_value_stop = rowcol_to_cell(last_line_pos, 7)
        sum_purchase_value_formula = 'SUM(%s:%s)' % (
            purchase_value_start, purchase_value_stop)  # @UnusedVariable
        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'totals'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.aml_cell_style_decimal)

        # Title "sổ tài sản cố định"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['left']
        cell_footer_style = xlwt.easyxf(cell_format)

        c_specs = [
            ('note1', 5, 0, 'text',
             '- Sổ này có ........ trang, đánh từ trang 01 dến trang ........', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)
        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('note1', 5, 0, 'text', '- Ngày mở sổ: .............', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)
        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = empty + [
            ('note1', 3, 0, 'text', 'Ngày ........ tháng ...... năm .........', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        # Add 2 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold']
        cell_footer_style = xlwt.easyxf(cell_format)

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 2, 0, 'text', 'Người ghi sổ', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 2, 0, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
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
            ('col1', 1, 0, 'text', '', None),
            ('col2', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col7', 1, 0, 'text', '', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)


asset_book_xls('report.asset.book.report',
               'account.asset.asset',
               parser=asset_book_xls_parser)
