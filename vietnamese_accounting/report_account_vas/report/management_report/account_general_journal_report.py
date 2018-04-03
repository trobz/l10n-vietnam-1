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
from openerp.addons.report_xls.report_xls import report_xls  # @UnresolvedImport
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from openerp.addons.report_base_vn.report import report_base_vn
from openerp.tools.safe_eval import safe_eval
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
import logging
_logger = logging.getLogger(__name__)


class account_general_journal_xls_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(account_general_journal_xls_parser, self).__init__(
                                                cr, uid, name, context=context)
        self.context = context
        self.acc_move = {}
        self.acc_move_line = {}
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_account_info': self.get_account_info,
            'get_report_param': self.get_report_param,
            'get_account_move_data': self.get_account_move_data,
            'get_lines_data': self.get_lines_data,
            'get_company': self.get_company,
            'get_date': self.get_date
        })

    def get_date(self):
        res = {
            'date_from_date': self.get_wizard_data().get(
                                            'date_from', datetime.today()),
            'date_to_date': self.get_wizard_data().get(
                                            'date_to', datetime.today())
        }
        res['date_from'] = datetime.strptime(
                                res['date_from_date'],
                                DEFAULT_SERVER_DATE_FORMAT
                            ).strftime('%d/%m/%Y')
        res['date_to'] = datetime.strptime(
                            res['date_to_date'],
                            DEFAULT_SERVER_DATE_FORMAT
                        ).strftime('%d/%m/%Y')
        return res

    def get_company(self):
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        invoice_serial_number = 'ISN has been omitted'  # res.company_id.invoice_serial_number
        address_list = [
            res.company_id.street or '',
            res.company_id.street2 or '',
            res.company_id.city or '',
            res.company_id.state_id and res.company_id.state_id.name or '',
            res.company_id.country_id and res.company_id.country_id.name or '',
        ]

        address_list = filter(None, address_list)
        address = ', '.join(address_list)
        vat = res.company_id.vat or ''
        return {
            'name': name,
            'address': address,
            'vat': vat,
            'invoice_serial_number': invoice_serial_number
        }

    def get_wizard_data(self):
        result = {}
        datas = self.localcontext['data']
        if datas:
            result = datas['form'] and datas['form']
        return result

    def get_report_param(self, param=''):
        report_param = self.pool.get('ir.config_parameter').get_param(
                self.cr, self.uid, 'account_general_journal_report_parameter')
        report_param = """{'format_form': 'Mẫu số: S05a – DN', 'form_formated_by_rule': '''(Ban hành theo QĐ số: 48/2006/QĐ- BTC ngày 14/9/2006 của Bộ trưởng BTC)'''}"""
        report_param = safe_eval(report_param)
        return report_param.get(param, '')

    def get_account_info(self):
        account_id = self.get_wizard_data().get('account', '') or False
        if account_id:
            res = self.pool.get('account.account').read(
                            self.cr, self.uid, account_id, ['code', 'name'])
            return res
        return {'code': '', 'name': ''}

    def get_account_move_data(self):
        """
        get account_move datas
        @param return:
            {
                account_id: name_of_entry
            }
        """
        date_info = self.get_date()

        params = {'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date']}
        SQL = """
            SELECT id AS invoice_id,
                   name AS serial,
                   date AS effective_date,
                   create_date AS date_created,
                   narration AS description,
                   to_char(date,'YYYYMMDD') AS effective_date_char
            FROM account_move
            WHERE date >= '%(date_start)s'
                AND date <= '%(date_end)s'
        """
        if self.get_wizard_data()['target_move'] == 'posted':
            SQL = SQL + "AND state = 'posted' ORDER BY date ASC, name"
        else:
            SQL = SQL + " ORDER BY date ASC, name"

        self.cr.execute(SQL % params)
        if not self.acc_move:
            for data in self.cr.dictfetchall():
                self.acc_move.update({data.get('invoice_id'): data})
        return self.acc_move

    def get_lines_data(self):
        """
        Get all journal entries and all journal items of this entry.
        @param return:
            {
                account_id: [ list_of_account_move_lines ]
            }

        """
        date_info = self.get_date()
        # self.get_account_move_data()
        account_move_ids = [-1, -1]
        if self.acc_move:
            account_move_ids = self.acc_move.keys()
        params = {'move_ids': tuple(account_move_ids + [-1, -1]),
                  'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date']}

        SQL = """
        SELECT  aml.move_id AS invoice_id,
                aml.name,
                account_account.code,
                aml.credit,
                aml.debit
        FROM account_move_line aml
            JOIN account_account
            ON aml.account_id = account_account.id
        WHERE date >= '%(date_start)s'
            AND date <= '%(date_end)s'
            AND move_id IN %(move_ids)s
         ORDER BY date ASC, name
        """
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()

        for move_line in data:
            if move_line.get('invoice_id', -1) not in self.acc_move_line:
                self.acc_move_line[move_line.get('invoice_id')] = [move_line]
            else:
                self.acc_move_line[move_line.get('invoice_id')].append(
                                                                    move_line)

        return self.acc_move_line


class account_general_journal_xls(report_xls, report_base_vn.Parser):

    def __init__(self, name, table, rml=False, parser=False, header=True,
                 store=False):
        super(account_general_journal_xls, self).__init__(
                                    name, table, rml, parser, header, store)

        # Cell Styles
        date_format = 'DD/MM/YYYY'
        decimal_format = '#,##0'
        _xs = self.xls_styles
        light_orange = 'pattern: pattern solid, fore_color light_yellow;'

        # move lines header
        aml_header_cell_format = _xs['borders_all'] + _xs['wrap'] +\
            _xs['middle'] + light_orange + _xs['center']
        self.aml_header_cell_style = xlwt.easyxf(aml_header_cell_format)
        self.aml_header_cell_style_left = xlwt.easyxf(
            _xs['borders_all'] + _xs['wrap'] + _xs['middle'] +
            light_orange + _xs['left'])
        self.aml_header_cell_style_date = xlwt.easyxf(
                            aml_header_cell_format, num_format_str=date_format)
        self.aml_header_cell_style_decimal = xlwt.easyxf(
                        aml_header_cell_format, num_format_str=decimal_format)

        # lines
        aml_cell_format = _xs['borders_all'] + _xs['wrap'] + _xs['middle']
        self.aml_cell_style_center = xlwt.easyxf(aml_cell_format +
                                                 _xs['center'] +
                                                 _xs['bold'])
        self.aml_cell_style_left = xlwt.easyxf(aml_cell_format + _xs['left'])
        self.aml_cell_style_date = xlwt.easyxf(aml_cell_format + _xs['left'] +
                                               _xs['center'],
                                               num_format_str=date_format)
        self.aml_cell_style_decimal = xlwt.easyxf(
                aml_cell_format + _xs['right'], num_format_str=decimal_format)
        self.aml_cell_style_decimal_bold = xlwt.easyxf(
                                aml_cell_format + _xs['right'] + _xs['bold'],
                                num_format_str=decimal_format)

        self.asset_xls_styles = {
            'normal': '',
            'bold': '',
            'underline': 'font: underline true;',
        }

        # normal
        self.cell_style_normal = xlwt.easyxf(self.asset_xls_styles['normal'] +
                                             _xs['borders_all'] + _xs['wrap'] +
                                             _xs['center'] + _xs['middle'])
        self.cell_style_normal_borderless = xlwt.easyxf(
                                self.asset_xls_styles['normal'] + _xs['wrap'] +
                                _xs['center'] + _xs['middle'])
        cell_total_style = xlwt.easyxf(
                                    _xs['wrap'] + _xs['center'] + _xs['bold'] +
                                    _xs['borders_all'] + _xs['middle'])

        # center
        self.cell_style_center = xlwt.easyxf(
                            _xs['center'] + _xs['borders_all'] + _xs['wrap'] +
                            _xs['middle'] + _xs['bold'])

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E', 'G', 'H']
        self.col_specs_template = {
            'A': {
                'move_lines': [1, 15, _render("line.get('date_created','') and 'date' or 'text'"), _render("datetime.strptime(line.get('date_created','')[:10],'%Y-%m-%d')"), None, self.aml_header_cell_style_date],
                'lines': [1, 15, 'text', None],
                'totals': [1, 0, 'text', None]},
            'B': {
                'move_lines': [1, 15, 'text', _render("line.get('serial','') or ''"), None, self.aml_header_cell_style_left],
                'lines': [1, 15, 'text', None],
                'totals': [1, 0, 'text', None]},
            'C': {
                'move_lines': [1, 15, _render("line.get('effective_date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('effective_date',None)[:10],'%Y-%m-%d')"), None, self.aml_header_cell_style_date],
                'lines': [1, 15, 'text', None],
                'totals': [1, 0, 'text', None]},
            'D': {
                'move_lines': [1, 50, 'text', _render("line.get('description','') or ''"), None, self.aml_header_cell_style_left],
                'lines': [1, 16, 'text', _render("move_line.get('name','') or ''"), None, self.aml_cell_style_left],
                'totals': [1, 16, 'text', 'Tổng Cộng', None, cell_total_style]},
            'E': {
                'move_lines': [1, 15, 'text', None, None, None],
                'lines': [1, 15, 'text', _render("move_line.get('code',None)")],
                'totals': [1, 0, 'text', None]},
            'G': {
                'move_lines': [1, 15, 'text', None, None, None],
                'lines': [1, 0, 'number', _render("move_line.get('debit','') or ''"), None, self.aml_cell_style_decimal],
                'totals': [1, 0, 'number', _render('total_debit'), None, self.aml_cell_style_decimal_bold]},
            'H': {
                'move_lines': [1, 15, 'text', None, None, None],
                'lines': [1, 0, 'number', _render("move_line.get('credit','') or ''"), None, self.aml_cell_style_decimal],
                'totals': [1, 0, 'number', _render('total_credit'), None, self.aml_cell_style_decimal_bold]},

        }

    def _get_header(self, _p, _xs, data, objects, wb, ws,
                    row_pos, report_name):
        """
        Display header title of the report
        """
        # set print header/footer
        ws.header_str = self.xls_headers['standard']
        ws.footer_str = self.xls_footers['standard']
        cell_address_format =\
            _xs['bold'] + _xs['wrap'] + _xs['left'] + _xs['top']
        cell_address_style = xlwt.easyxf(cell_address_format)
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold'] + _xs['top']
        cell_footer_style = xlwt.easyxf(cell_format)

        # Title 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 500
        c_specs = [
            ('company_name', 3, 0, 'text',
             u'Đơn vị: %s' % _p.get_company()['name'], '', cell_address_style),
            ('form_serial', 4, 0, 'text', _p.get_report_param('format_form'),
             '', cell_footer_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal_borderless)

        # Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 800
        c_specs = [
            ('company_name', 3, 0, 'text',
             u'Địa chỉ: %s' % _p.get_company()['address'], '', cell_address_style),
            ('form_serial', 4, 0, 'text',
             _p.get_report_param('form_formated_by_rule'), '', self.cell_style_normal_borderless)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.cell_style_normal_borderless)

        # Title 3
        c_specs = [
            ('company_name', 3, 0, 'text',
             u'MST: %s' % _p.get_company()['vat'] or '', '', cell_address_style),
            ('empty1', 1, 0, 'text', ''),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal_borderless)

        # Title "SỔ NHẬT KÝ CHUNG"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_format = _xs['bold'] + _xs['wrap'] +\
            _xs['center'] + 'font: height 360;'
        cell_title_style = xlwt.easyxf(cell_format)
        c_specs = [
            ('payment_journal', 7, 0, 'text', report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal_borderless)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['italic'] + _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_title_style = xlwt.easyxf(cell_format)
        c_specs = [
            ('asset_book', 7, 0, 'text',
             u'Từ %s đến %s' % (_p.get_date().get('date_from', '.......'),
                                _p.get_date().get('date_to', '.......')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=cell_title_style)

        # Add 1 empty line
        c_specs = [
            ('col1', 1, 15, 'text', '', None),
            ('col2', 1, 15, 'text', '', None),
            ('col3', 1, 15, 'text', '', None),
            ('col4', 1, 50, 'text', '', None),
            ('col5', 1, 15, 'text', '', None),
            ('col6', 1, 15, 'text', '', None),
            ('col7', 1, 15, 'text', '', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal_borderless)

        # Header Title 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 400
        row_title_body_pos = row_pos
        c_specs = [
            ('col1', 1, 30, 'text', u'Ngày tháng ghi sổ', None, self.cell_style_center),
            ('col2', 2, 0, 'text', u'Chứng Từ', None, self.cell_style_center),
            ('col3', 1, 0, 'text', u'Diễn giải', None, self.cell_style_center),
            ('col4', 1, 0, 'text', u'Tài khoản', None, self.cell_style_center),
            ('col5', 2, 0, 'text', u'Số tiền', None, self.cell_style_center),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal)

        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 400
        c_specs = [
            ('col1', 1, 0, 'text', u'', None, self.cell_style_center),
            ('col2', 1, 0, 'text', u'Số hiệu', None, self.cell_style_center),
            ('col3', 1, 0, 'text', u'Ngày tháng', None, self.cell_style_center),
            ('col4', 1, 0, 'text', u'', None, self.cell_style_center),
            ('col5', 1, 0, 'text', u'', None, self.cell_style_center),
            ('col6', 1, 0, 'text', u'Nợ', None, self.cell_style_center),
            ('col7', 1, 0, 'text', u'Có', None, self.cell_style_center),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 0, 0,
                       u'Ngày tháng ghi sổ', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 3, 3,
                       u'Diễn giải', self.aml_cell_style_center)
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 4, 4,
                       u'Tài khoản', self.aml_cell_style_center)

        return row_pos

    def _get_worksheet(self, _p, _xs, data, objects, wb, report_name, count=0):
        """
        @summary: get new worksheet from workbook, reset current row position
                    in the new worksheet
        @return: new worksheet, new row_pos
        """
        report_name = count and (
                            report_name[:31] + str(count)) or report_name[:31]
        ws = wb.add_sheet(report_name, cell_overwrite_ok=True)
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_width_to_pages = 1
        row_pos = 0

        return ws, row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        report_name = u"SỔ NHẬT KÝ CHUNG"
        count = 1
        ws, row_pos = self._get_worksheet(_p, _xs, data, objects,
                                          wb, report_name)

        MAX_ROW = 65000

        # display header/footer
        row_pos = self._get_header(_p, _xs, data, objects, wb, ws,
                                   row_pos, report_name)

        # account move lines
        move_data = _p.get_account_move_data()
        # only get the values of dict
        move_data = move_data.values()
        # sorted by id
        move_data = sorted(move_data, key=lambda x: x.get('effective_date_char'))
        move_line_data = _p.get_lines_data()
        total_debit = 0
        total_credit = 0
        for line in move_data:
            if row_pos > MAX_ROW:
                ws.flush_row_data()
                ws, row_pos = self._get_worksheet(_p, _xs, data, objects, wb,
                                                  report_name, count)
                row_pos = self._get_header(_p, _xs, data, objects, wb, ws,
                                           row_pos, report_name)
                count += 1

            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 400
            c_specs = map(lambda x: self.render(x, self.col_specs_template,
                                                'move_lines'),
                          self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(ws, row_pos, row_data,
                                         row_style=self.aml_header_cell_style,
                                         set_column_size=True)

            for move_line in move_line_data.get(line.get('invoice_id'), []):
                ws.row(row_pos).height_mismatch = True
                ws.row(row_pos).height = 400
                c_specs = map(lambda x: self.render(x, self.col_specs_template,
                                                    'lines'), self.wanted_list)
                row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
                row_pos = self.xls_write_row(ws, row_pos, row_data,
                                             row_style=self.cell_style_normal)
                total_debit += move_line.get('debit', 0)
                total_credit += move_line.get('credit', 0)

        # Totals
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 400
        c_specs = map(lambda x: self.render(x, self.col_specs_template,
                                            'totals'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.aml_cell_style_decimal)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.cell_style_normal_borderless)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(5)]
        c_specs = empty + [
            ('note1', 2, 0, 'text', u'Ngày .... tháng .... năm ....', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['bold']
        cell_footer_style = xlwt.easyxf(cell_format)

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', u'Người ghi sổ', None),
            ('col3', 1, 0, 'text', u'', None),
            ('col6', 1, 16, 'text', u'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', u'', None),
            ('col9', 2, 0, 'text', u'Giám đốc', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=cell_footer_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['italic']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col2', 2, 0, 'text', u'(Ký, họ tên)', None),
            ('col5', 1, 0, 'text', u'', None),
            ('col6', 1, 16, 'text', u'(Ký, họ tên)', None),
            ('col8', 1, 0, 'text', u'', None),
            ('col9', 2, 0, 'text', u'(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=cell_footer_style)

account_general_journal_xls('report.vas_account_general_journal',
                            'account.move',
                            parser=account_general_journal_xls_parser)

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
