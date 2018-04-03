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


from . import account_profit_and_loss_parser as parser
import xlwt
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from .. import report_xls_utils
import time
from openerp.addons.report_base_vn.report import report_base_vn
from datetime import datetime
from datetime import timedelta
# import ast
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT


class Parser(parser.Parser):

    def __init__(self, cr, uid, name, context=None):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'account_profit_and_loss_report_xls'
        self.localcontext.update({
        })


class account_profit_and_loss_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True,
                 store=False):
        super(account_profit_and_loss_xls, self).__init__(
            name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

    def generate_xls_ws(self, _p, _xs, data, objects, wb, ws, row_pos,
                        report_name):
        """
        @return: row_pos: position of at the end of generatioon header
        """
        num_format_str = '#,##0 ;(#,##0)'
        cell_address_style = self.get_cell_style(
            ['bold', 'wrap', 'left', 'top'])
        cell_footer_style = self.get_cell_style(
            ['bold', 'wrap', 'center', 'top'])
        cell_body_style = self.get_cell_style(
            ['bold', 'wrap', 'center', 'top', 'borders_all'])
        cell_normal_style = self.get_cell_style(
            ['bold', 'wrap', 'left', 'top', 'borders_all'])
        cell_number_style = self.get_cell_style(
            ['wrap', 'center', 'top', 'borders_all'])
        cell_italic_style = self.get_cell_style(
            ['wrap', 'left', 'top', 'italic'])
        cell_result_style = self.get_cell_style(
            ['wrap', 'right', 'top', 'borders_all'], num_format_str)
        cell_result_style_bold = self.get_cell_style(
            ['bold', 'wrap', 'right', 'top', 'borders_all'], num_format_str)
        # Title address 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('company_detail', 2, 35, 'text', u'Đơn vị: %s' %
             _p.get_company()['name'] or '', None, cell_address_style),
            ('col_no', 1, 10, 'text', '', None, cell_address_style),
            ('form_serial', 3, 60, 'text',
             u'''Mẫu số B02 – DN''', None, self.normal_style_bold)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style, set_column_size=True)

        # Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 500
        c_specs = [
            ('company_detail', 2, 0, 'text', u'Địa chỉ: %s' %
             _p.get_company()['address'] or '', '', cell_address_style),
            ('col_no', 1, 0, 'text', '', None, cell_address_style),
            ('form_serial', 3, 27, 'text',
             u'''(Ban hành theo QĐ số 15/2006/QĐ-BTC, Ngày 20/03/2006 của Bộ trưởng BTC)''',
             '', self.normal_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_title_style = self.get_cell_style(
            ['bold', 'wrap', 'center', 'middle', 'fontsize_350'])

        c_specs = [
            ('profit_and_loss', 6, 0, 'text',
             report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('asset_book', 6, 0, 'text', 'Từ %s đến %s' %
             (_p.get_date(data)['date_from'], _p.get_date(data)['date_to']))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 2 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        c_specs = [
            ('col1', 1, 17, 'text', '', None),
            ('col2', 1, 18, 'text', '', None),
            ('col3', 1, 10, 'text', '', None),
            ('col4', 1, 20, 'text', '', None),
            ('col5', 2, 40, 'text', u'Đơn vị tính: %s' %
             _p.get_company()['currency'], None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style,
            set_column_size=True)

        # Header title 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 2, 35, 'text', u'Chỉ tiêu', None, cell_body_style),
            ('col2', 1, 10, 'text', u'Mã số', None, cell_body_style),
            ('col3', 1, 20, 'text', u'Thuyết minh',
             None, cell_body_style),
            ('col4', 1, 20, 'text', u'Kỳ báo cáo',
             None, cell_body_style),
            ('col5', 1, 20, 'text', u'Cùng kì năm trước',
             None, cell_body_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style,
            set_column_size=True)
        # Line 1
        c_specs = [
            ('col1', 2, 0, 'text', u'1. Doanh thu bán hàng và cung cấp dịch vụ',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '01', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(1, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(1, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 2
        c_specs = [
            ('col1', 2, 0, 'text', u'2. Các khoản giảm trừ',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '02', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(2, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(2, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 3
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('col1', 2, 0, 'text', u'''3. Doanh thu bán hàng thuần và cung cấp dịch vụ
    (10 = 01 - 02)''',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '10', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '', 'E9 - E10', cell_result_style_bold),
            ('col5', 1, 0, 'number', '', 'F9 - F10', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 4
        c_specs = [
            ('col1', 2, 0, 'text', u'4. Giá vốn hàng bán',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '11', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(11, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(11, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 5
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('col1', 2, 0, 'text', u'''5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ
    (20 = 10 - 11)''',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '20', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '', 'E11-E12', cell_result_style_bold),
            ('col5', 1, 0, 'number', '', 'F11-F12', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 6
        c_specs = [
            ('col1', 2, 0, 'text', u'6. Doanh thu hoạt động tài chính',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '21', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(21, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(21, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 7
        c_specs = [
            ('col1', 2, 0, 'text', u' 7. Chi phí tài chính',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '22', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(22, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(22, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Chi phi vay lai
        c_specs = [
            ('col1', 2, 0, 'text', u'Trong đó: Chi phí lãi vay',
             None, cell_italic_style),
            ('col2', 1, 0, 'text', '23', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(23, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(23, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 8
        c_specs = [
            ('col1', 2, 0, 'text', u'8. Chi phí bán hàng',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '24', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(24, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(24, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 9
        c_specs = [
            ('col1', 2, 0, 'text', u'9. Chi phí quản lý doanh nghiệp',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '25', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(25, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(25, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 10
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('col1', 2, 0, 'text', u'''10. Lợi nhuận thuần từ hoạt động kinh doanh
    (30 = 20 + (21 - 22) - (24 + 25))''',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '30', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '',
             'E13+(E14-E15)-(E17+E18)', cell_result_style_bold),
            ('col5', 1, 0, 'number', '',
             'F13+(F14-F15)-(F17+F18)', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 11
        c_specs = [
            ('col1', 2, 0, 'text', u'11. Thu nhập khác',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '31', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(31, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(31, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 12
        c_specs = [
            ('col1', 2, 0, 'text', u'12. Chi phí khác',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '32', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(32, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(32, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 13
        c_specs = [
            ('col1', 2, 0, 'text', u'13. Lợi nhuận khác (40 = 31 -32)',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '40', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '', 'E20-E21', cell_result_style_bold),
            ('col5', 1, 0, 'number', '', 'F20-F21', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 14
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('col1', 2, 0, 'text', u'''14. Tổng lợi nhuận kế toán trước thuế
    (50 = 30 + 40)''',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '50', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '', 'E19+E22', cell_result_style_bold),
            ('col5', 1, 0, 'number', '', 'F19+F22', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 15
        c_specs = [
            ('col1', 2, 0, 'text', u'15. Chi phí thuế TNDN hiện hành',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '51', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(51, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(51, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 16
        c_specs = [
            ('col1', 2, 0, 'text', u'16. Chi phí thuế TNDN hoãn lại',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '52', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', _p.get_result(52, 'now'),
             None, cell_result_style),
            ('col5', 1, 0, 'number', _p.get_result(52, 'last'),
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 17
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        c_specs = [
            ('col1', 2, 0, 'text', u'''17. Lợi nhuận sau thuế thu nhập doanh nghiệp
    (60 = 50 - 51 - 52)''',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '60', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'number', '',
             'E23-E24-E25', cell_result_style_bold),
            ('col5', 1, 0, 'number', '',
             'F23-F24-F25', cell_result_style_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line 18
        c_specs = [
            ('col1', 2, 0, 'text', u'18. Lãi cơ bản trên cổ phiếu (*)',
             None, cell_normal_style),
            ('col2', 1, 0, 'text', '70', None, cell_number_style),
            ('col3', 1, 0, 'text', '',
             None, cell_normal_style),
            ('col4', 1, 0, 'text', '',
             None, cell_result_style),
            ('col5', 1, 0, 'text', '',
             None, cell_result_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line_footer_1
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 2, 0, 'text', u''' Lập, ngày ... tháng ... năm ... ''',
             None, self.normal_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line_footer_2
        c_specs = [
            ('col1', 1, 0, 'text', u'Người lập biểu',
             None, cell_footer_style),
            ('col2', 1, 0, 'text', '', None, cell_footer_style),
            ('col3', 2, 0, 'text', u'Kế toán trưởng',
             None, cell_footer_style),
            ('col4', 2, 0, 'text', u'Giám đốc', None, cell_footer_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Line_footer_3
        c_specs = [
            ('col1', 1, 0, 'text', u'(Ký, họ tên)', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 2, 0, 'text', u'(Ký, họ tên)', None),
            ('col4', 2, 0, 'text', u'(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)
        # Add 4 empty line
        for _n in range(4):
            c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data, row_style=self.normal_style)

        c_specs = [
            ('col1', 3, 0, 'text', u'Ghi chú: (*) Chỉ tiêu này chỉ áp dụng đối với công ty cổ phần',
             None, self.normal_style),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', u'', None),
            ('col4', 1, 0, 'text', u'', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)
        return row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        MAX_ROW = 65500
        count = 1
        report_name = u"BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH"

        # set print sheet
        ws = super(account_profit_and_loss_xls, self).generate_xls_report(
            _p, _xs, data, objects, wb, report_name)
        row_pos = 0

        row_pos = self.generate_xls_ws(
            _p, _xs, data, objects, wb, ws, row_pos, report_name)
        ws.show_grid = False
        ws.portrait = True


account_profit_and_loss_xls('report.account_profit_and_loss_report_xls',
                            'account.account', parser=Parser)


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
