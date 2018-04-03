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


from datetime import datetime
from openerp.addons.report_base_vn.report import report_base_vn
from openerp.tools.safe_eval import safe_eval
from openerp.addons.report_xls.report_xls import report_xls
import xlwt
from openerp.addons.report_xls.utils import _render
from openerp.addons.report_xls.utils import rowcol_to_cell
from account_balance_sheet_report import balance_sheet_report_parser
import inspect

"""
if type = 1.1 :  bold, align center
   type = 1.2 : not bold, not align center
   type = 2.1.1 : bold,  style of each column is different, formular
   type = 2.1.2 : bold,  style of each column is different, function
   type = 2.2 : not bold, tyle of each column is different, function
"""
define_template = [
    ['Chỉ tiêu', 'Mã số', 'Thuyết Minh',
        'Số cuối năm (3)', 'Số đầu năm (3)', '1.1'],
    [1, 2, 3, 4, 5, '1.2'],
    ['TÀI SẢN', ' ', ' ', ' ', ' ', '1.1'],
    ['A - TÀI SẢN NGẮN HẠN (100=110+120+130+140+150)', '100', ' ',
     'SUM(D12,D15,D18,D25,D28)', 'SUM(E12,E15,E18,E25,E28)', '2.1.1'],
    ['I. Tiền và các khoản tương đương tiền', '110', ' ',
     'get_total(110)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1.Tiền ', '111', 'V.01',
     "get_every_line('111')", 'return_beginning_balance()', '2.2'],
    [' 2. Các khoản tương đương tiền ', '112', ' ',
     "get_every_line('112')", 'return_beginning_balance()', '2.2'],
    ['II. Các khoản đầu tư tài chính ngắn hạn', '120', 'V.02',
     "get_total(120)", 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Chứng khoán kinh doanh', '121', ' ',
     'get_every_line(121)', 'return_beginning_balance()', '2.2'],
    [' 2. Dự phòng giảm giá chứng khoán kinh doanh', '122', ' ',
     'get_every_line(122)', 'return_beginning_balance()', '2.2'],
    [' 3. Đầu tư nắm giữ đến ngày đáo hạn', '123', ' ',
     'get_every_line(123)', 'return_beginning_balance()', '2.2'],
    ['III. Các khoản phải thu ngắn hạn', '130', ' ',
     'get_total(130)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Phải thu ngắn hạn của khách hàng', '131', ' ',
     'get_every_line(131)', 'return_beginning_balance()', '2.2'],
    [' 2. Trả trước cho người bán ngắn hạn', '132', ' ',
     'get_every_line(132)', 'return_beginning_balance()', '2.2'],
    [' 3. Phải thu nội bộ ngắn hạn', '133', ' ',
     'get_every_line(133)', 'return_beginning_balance()', '2.2'],
    [' 4. Phải thu theo tiến độ kế hoạch hợp đồng xây dựng', '134',
     ' ', 'get_every_line(134)', 'return_beginning_balance()', '2.2'],
    [' 5. Phải thu về cho vay ngắn hạn', '135', 'V.03',
     'get_every_line(135)', 'return_beginning_balance()', '2.2'],
    [' 6. Phải thu ngắn hạn khác', '136', 'V.03',
     'get_every_line(136)', 'return_beginning_balance()', '2.2'],
    [' 7. Dự phòng phải thu ngắn hạn khó đòi (*)', '137', ' ',
     'get_every_line(137)', 'return_beginning_balance()', '2.2'],
    [' 7. Tài sản thiếu chờ xử lý', '139', ' ',
     'get_every_line(139)', 'return_beginning_balance()', '2.2'],
    ['IV. Hàng tồn kho', '140', ' ',
     'get_total(140)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Hàng tồn kho', '141', 'V.04',
     'get_every_line(141)', 'return_beginning_balance()', '2.2'],
    [' 2. Dự phòng giảm giá hàng tồn kho (*)', '149', ' ',
     'get_every_line(149)', 'return_beginning_balance()', '2.2'],
    ['V. Tài sản ngắn hạn khác', '150', ' ',
     'get_total(150)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Chi phí trả trước ngắn hạn', '151', ' ',
     'get_every_line(151)', 'return_beginning_balance()', '2.2'],
    [' 2. Thuế GTGT được khấu trừ', '152', ' ',
     'get_every_line(152)', 'return_beginning_balance()', '2.2'],
    [' 3. Thuế và các khoản khác phải thu Nhà nước', '153', 'V.05',
     'get_every_line(153)', 'return_beginning_balance()', '2.2'],
    [' 4.  Giao dịch mua bán lại trái phiếu Chính Phủ', '154', ' ',
     'get_every_line(154)', 'return_beginning_balance()', '2.2'],
    [' 5. Tài sản ngắn hạn khác', '155', ' ',
     'get_every_line(155)', 'return_beginning_balance()', '2.2'],
    ['B - TÀI SẢN DÀI HẠN (200 = 210 + 220 + 240 + 250 + 260)', '200',
     ' ', 'SUM(D34,D40,D51,D54,D59)', 'SUM(E34,E40,E51,E54,E59)', '2.1.1'],
    ['I- Các khoản phải thu dài hạn', '210', ' ',
     'get_total(210)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Phải thu dài hạn của khách hàng', '211', ' ',
     'get_every_line(211)', 'return_beginning_balance()', '2.2'],
    [' 2. Trả trước cho người bán dài hạn', '212', ' ',
     'get_every_line(212)', 'return_beginning_balance()', '2.2'],
    [' 3. Vốn kinh doanh ở đơn vị trực thuộc', '213', ' ',
     'get_every_line(213)', 'return_beginning_balance()', '2.2'],
    [' 4. Phải thu nội bộ dài hạn ', '214', '',
     'get_every_line(213)', 'return_beginning_balance()', '2.2'],
    [' 5. Phải thu về cho vay dài hạn', '215', '',
     'get_every_line(215)', 'return_beginning_balance()', '2.2'],
    [' 5. Phải thu dài hạn khác ', '216', '',
     'get_every_line(216)', 'return_beginning_balance()', '2.2'],
    [' 6. Dự phòng phải thu dài hạn khó đòi (*) ', '219', ' ',
     'get_every_line(219)', 'return_beginning_balance()', '2.2'],
    ['II. Tài sản cố định', '220', ' ',
     'SUM(D41,D44,D47,D50)', 'SUM(E41,E44,E47,E50)', '2.1.1'],
    [' 1. Tài sản cố định hữu hình', '221', 'V.08',
     'get_total(221)', 'return_beginning_balance_total()', '2.2'],
    ['  - Nguyên giá', '222', ' ',
     'get_every_line(222)', 'return_beginning_balance()', '2.2'],
    ['  - Giá trị hao mòn luỹ kế (*)', '223', ' ',
     'get_every_line(223)', 'return_beginning_balance()', '2.2'],
    [' 2. Tài sản cố định thuê tài chính', '224', '',
     'get_total(224)', 'return_beginning_balance_total()', '2.2'],
    ['  - Nguyên giá', '225', ' ',
     'get_every_line(225)', 'return_beginning_balance()', '2.2'],
    ['  - Giá trị hao mòn luỹ kế (*)', '226', ' ',
     'get_every_line(226)', 'return_beginning_balance()', '2.2'],
    [' 3. Tài sản cố định vô hình', '227', 'V.10',
     'get_total(227)', 'return_beginning_balance_total()', '2.2'],
    ['  - Nguyên giá', '228', ' ',
     'get_every_line(228)', 'return_beginning_balance()', '2.2'],
    ['  - Giá trị hao mòn luỹ kế (*)', '229', ' ',
     'get_every_line(229)', 'return_beginning_balance()', '2.2'],
    ['III. Bất động sản đầu tư', '230', '',
     'get_total(230)', 'return_beginning_balance_total()', '2.1.2'],
    ['  - Nguyên giá', '231', ' ',
     'get_every_line(231)', 'return_beginning_balance()', '2.2'],
    ['  - Giá trị hao mòn luỹ kế (*)', '232', ' ',
     'get_every_line(232)', 'return_beginning_balance()', '2.2'],
    ['IV.  Tài sản dở dang dài hạn', '240', ' ',
     'get_total(240)', 'return_beginning_balance_total()', '2.1.2'],
    ['1. Chi phí sản xuất, kinh doanh dở dang dài hạn', '241', ' ',
     'get_every_line(241)', 'return_beginning_balance()', '2.2'],
    ['2. Chi phí xây dựng cơ bản dở dang', '242', ' ',
     'get_every_line(242)', 'return_beginning_balance()', '2.2'],
    ['V. Các khoản đầu tư tài chính dài hạn', '250', ' ',
     'get_total(250)', 'return_beginning_balance_total()', '2.1.2'],
    ['1. Đầu tư vào công ty con ', '251', ' ',
     'get_every_line(251)', 'return_beginning_balance()', '2.2'],
    ['2. Đầu tư vào công ty liên kết, liên doanh ', '252', ' ',
     'get_every_line(252)', 'return_beginning_balance()', '2.2'],
    ['3. Đầu tư góp vốn vào đơn vị khác', '253', ' ',
     'get_every_line(253)', 'return_beginning_balance()', '2.2'],
    ['4. Dự phòng giảm giá đầu tư tài chính dài hạn (*) ', '254',
     ' ', 'get_every_line(254)', 'return_beginning_balance()', '2.2'],
    ['5. Đầu tư nắm giữ đến ngày đáo hạn', '255',
     ' ', 'get_every_line(255)', 'return_beginning_balance()', '2.2'],
    ['VI. Tài sản dài hạn khác', '260', ' ',
     'get_total(260)', 'return_beginning_balance_total()', '2.1.2'],
    ['1. Chi phí trả trước dài hạn ', '261', 'V.14',
     'get_every_line(261)', 'return_beginning_balance()', '2.2'],
    ['2. Tài sản thuế thu nhập hoãn lại ', '262', '',
     'get_every_line(262)', 'return_beginning_balance()', '2.2'],
    ['3. Thiết bị, vật tư, phụ tùng thay thế dài hạn', '263', '',
     'get_every_line(263)', 'return_beginning_balance()', '2.2'],
    ['4. Tài sản dài hạn khác ', '268', ' ',
     'get_every_line(268)', 'return_beginning_balance()', '2.2'],
    ['TỔNG CỘNG TÀI SẢN (270 = 100 + 200)', '270',
     ' ', 'SUM(D11,D33)', 'SUM(E11,E33)', '2.1.1'],
    [' ', ' ', ' ', ' ', ' ', '1.1'],
    ['NGUỒN VỐN', ' ', ' ', ' ', ' ', '1.1'],
    ['A. NỢ PHẢI TRẢ (300 = 310 + 330)', '300',
     ' ', 'SUM(D67,D78)', 'SUM(E67,E78)', '2.1.1'],
    ['I. Nợ ngắn hạn', '310', ' ',
     'get_total(310)', 'return_beginning_balance_total()', '2.1.2'],
    ['1. Phải trả người bán ngắn hạn', '311', '',
     'get_every_line(311)', 'return_beginning_balance()', '2.2'],
    ['2. Người mua trả tiền trước ngắn hạn', '312', ' ',
     'get_every_line(312)', 'return_beginning_balance()', '2.2'],
    ['3. Thuế và các khoản phải nộp nhà nước', '313', ' ',
     'get_every_line(313)', 'return_beginning_balance()', '2.2'],
    ['4. Phải trả người lao động', '314', '',
     'get_every_line(314)', 'return_beginning_balance()', '2.2'],
    ['5. Chi phí phải trả ngắn hạn', '315', ' ',
     'get_every_line(315)', 'return_beginning_balance()', '2.2'],
    ['6. Phải trả nội bộ ngắn hạn', '316', '',
     'get_every_line(316)', 'return_beginning_balance()', '2.2'],
    ['7. Phải trả theo tiến độ kế hoạch hợp đồng xây dựng', '317', ' ',
     'get_every_line(317)', 'return_beginning_balance()', '2.2'],
    ['8. Doanh thu chưa thực hiện ngắn hạn', '318',
     ' ', 'get_every_line(318)', 'return_beginning_balance()', '2.2'],
    ['9. Phải trả ngắn hạn khác', '319', '',
     'get_every_line(319)', 'return_beginning_balance()', '2.2'],
    ['10. Vay và nợ thuê tài chính ngắn hạn', '320', ' ',
     'get_every_line(320)', 'return_beginning_balance()', '2.2'],
    ['11. Dự phòng phải trả ngắn hạn', '321', ' ',
     'get_every_line(321)', 'return_beginning_balance()', '2.2'],
    ['12. Quỹ khen thưởng, phúc lợi', '322', ' ',
     'get_every_line(322)', 'return_beginning_balance()', '2.2'],
    ['13. Quỹ bình ổn giá', '323', ' ',
     'get_every_line(323)', 'return_beginning_balance()', '2.2'],
    ['14. Giao dịch mua bán lại trái phiếu Chính Phủ', '324', ' ',
     'get_every_line(324)', 'return_beginning_balance()', '2.2'],
    ['II. Nợ dài hạn', '330', ' ',
     'get_total(330)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Phải trả người bán dài hạn', '331', ' ',
     'get_every_line(331)', 'return_beginning_balance()', '2.2'],
    [' 2. Người mua trả tiền trước dài hạn', '332', '',
     'get_every_line(332)', 'return_beginning_balance()', '2.2'],
    [' 3. Chi phí phải trả dài hạn', '333', ' ',
     'get_every_line(333)', 'return_beginning_balance()', '2.2'],
    [' 4. Phải trả nội bộ về vốn kinh doanh', '334', 'V.20',
     'get_every_line(334)', 'return_beginning_balance()', '2.2'],
    [' 5. Phải trả nội bộ dài hạn', '335', 'V.21',
     'get_every_line(335)', 'return_beginning_balance()', '2.2'],
    [' 6. Doanh thu chưa thực hiện dài hạn', '336', ' ',
     'get_every_line(336)', 'return_beginning_balance()', '2.2'],
    [' 7. Phải trả dài hạn khác', '337', ' ',
     'get_every_line(337)', 'return_beginning_balance()', '2.2'],
    [' 8. Vay và nợ thuê tài chính dài hạn', '338', ' ',
     'get_every_line(338)', 'return_beginning_balance()', '2.2'],
    [' 9. Trái phiếu chuyển đổi', '339', ' ',
     'get_every_line(339)', 'return_beginning_balance()', '2.2'],
    [' 10. Cổ phiếu ưu đãi', '340', ' ',
     'get_every_line(340)', 'return_beginning_balance()', '2.2'],
    [' 11.  Thuế thu nhập hoãn lại phải trả', '341', ' ',
     'get_every_line(341)', 'return_beginning_balance()', '2.2'],
    [' 12. Dự phòng phải trả dài hạn', '342', ' ',
     'get_every_line(342)', 'return_beginning_balance()', '2.2'],
    [' 13. Quỹ phát triển khoa học và công nghệ', '343', ' ',
     'get_every_line(343)', 'return_beginning_balance()', '2.2'],

    ['B - VỐN CHỦ SỞ HỮU (400 = 410 + 430)', '400',
     ' ', 'SUM(D87,D99)', 'SUM(E87,E99)', '2.1.1'],
    ['I. Vốn chủ sở hữu', '410', '',
     'get_total(410)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Vốn góp của chủ sở hữu', '411', ' ',
     'get_every_line(411)', 'return_beginning_balance()', '2.2'],
    [' 2. Thặng dư vốn cổ phần', '412', ' ',
     'get_every_line(412)', 'return_beginning_balance()', '2.2'],
    [' 3. Quyền chọn chuyển đổi trái phiếu', '413', ' ',
     'get_every_line(413)', 'return_beginning_balance()', '2.2'],
    [' 4. Vốn khác của chủ sở hữu ', '414', ' ',
     'get_every_line(414)', 'return_beginning_balance()', '2.2'],
    [' 5. Cổ phiếu quỹ (*) ', '415', ' ', 'get_every_line(415)',
     'return_beginning_balance()', '2.2'],
    [' 6. Chênh lệch đánh giá lại tài sản ', '416', ' ',
     'get_every_line(416)', 'return_beginning_balance()', '2.2'],
    [' 7. Chênh lệch tỷ giá hối đoái', '417', ' ',
     'get_every_line(417)', 'return_beginning_balance()', '2.2'],
    [' 8. Quỹ đầu tư phát triển', '418', ' ',
     'get_every_line(418)', 'return_beginning_balance()', '2.2'],
    [' 9. Quỹ hỗ trợ sắp xếp doanh nghiệp', '419', ' ',
     'get_every_line(419)', 'return_beginning_balance()', '2.2'],
    [' 10. Quỹ khác thuộc vốn chủ sở hữu', '420', ' ',
     'get_every_line(420)', 'return_beginning_balance()', '2.2'],
    [' 11. Lợi nhuận sau thuế chưa phân phối', '421', ' ',
     'get_every_line(421)', 'return_beginning_balance()', '2.2'],
    ['    - LNST chưa phân phối lũy kế đến cuối kỳ trước', '421a', ' ',
     'get_every_line(421)', 'return_beginning_balance()', '2.2'],
    ['    - LNST chưa phân phối kỳ này', '421b', ' ',
     'get_every_line(421)', 'return_beginning_balance()', '2.2'],
    [' 12. Nguồn vốn đầu tư XDCB', '422', ' ',
     'get_every_line(422)', 'return_beginning_balance()', '2.2'],
    ['II. Nguồn kinh phí và quỹ khác', '430', ' ',
     'get_total(430)', 'return_beginning_balance_total()', '2.1.2'],
    [' 1. Quỹ khen thưởng, phúc lợi', '431', ' ',
     'get_every_line(431)', 'return_beginning_balance()', '2.2'],
    [' 2. Nguồn kinh phí ', '432', 'V.23',
     'get_every_line(432)', 'return_beginning_balance()', '2.2'],
    [' 3. Nguồn kinh phí đã hình thành TSCĐ', '433', ' ',
     'get_every_line(433)', 'return_beginning_balance()', '2.2'],
    ['TỔNG CỘNG NGUỒN VỐN (440 = 300 + 400)', '440',
     ' ', 'SUM(D66,D86)', 'SUM(E66,E86)', '2.1.1'],
    [' ', ' ', ' ', ' ', ' ', '1.1'],
    [' ', ' ', ' ', ' ', ' ', '1.1'],
]


class balance_sheet_report(report_xls):

    def __init__(self, name, table, rml=False, parser=False, header=True,
                 store=False):
        super(balance_sheet_report, self).__init__(
            name, table, rml, parser, header, store)
        _xs = self.xls_styles
        font_name = 'font: name Arial ;'
        self.rh_cell_style_header_1 = xlwt.easyxf(_xs['wrap'] +
                                                  _xs['left'] +
                                                  _xs['bold'] +
                                                  font_name)
        self.rh_cell_style_header_1_2 = xlwt.easyxf(_xs['wrap'] +
                                                    _xs['right'] +
                                                    _xs['italic'] +
                                                    font_name)

        self.rh_cell_style_header_1_1 = xlwt.easyxf(_xs['wrap'] +
                                                    _xs['left'] +
                                                    font_name)

        self.rh_cell_style_header_2 = xlwt.easyxf(_xs['wrap'] +
                                                  _xs['center'] +
                                                  _xs['bold'] +
                                                  font_name)

        self.rh_cell_style_header_3 = xlwt.easyxf(_xs['wrap'] +
                                                  _xs['center'] +
                                                  font_name)
        self.rh_cell_style_header_3_2 = xlwt.easyxf(_xs['wrap'] +
                                                    _xs['top'] +
                                                    _xs['center'] +
                                                    font_name)
        self.rh_cell_style_header_3_1 = xlwt.easyxf(_xs['wrap'] +
                                                    _xs['center'] +
                                                    _xs['italic'] +
                                                    font_name)
        self.rh_cell_style_header_4 = \
            xlwt.easyxf(_xs['wrap'] +
                        _xs['center'] +
                        _xs['bold'] +
                        'font: bold true, height 300;' +
                        font_name)

        self.rh_cell_style_header_5 = \
            xlwt.easyxf('borders: ' +
                        'left thin, right thin, top thin, bottom thin;' +
                        _xs['wrap'] +
                        _xs['center'] +
                        _xs['bold'] +
                        'font: bold true, height 300;' +
                        font_name)

    def get_cell_style(self, styles, cell_style_format=None):
        """
        Get default style for the cell
        @param styles: list of style which want to use
        @rtype: rowstyle
        """
        res_style = ''
        for style in styles:
            res_style += self.xls_styles[style]
        return xlwt.easyxf(res_style, num_format_str=cell_style_format)

    def generate_xls_report(self, _p, _xs, data, objects, wb):
        def render(str_code):
            caller_space = inspect.currentframe().f_back.f_back.f_locals
            localcontext = self.parser_instance.localcontext
            render_space = caller_space
            render_space.update(localcontext)
            result = eval(str_code, render_space)
            return result

        report_name = "Balance Sheet"
        ws = wb.add_sheet(report_name)
        row_pos = 0
        c_specs = [('heading_1', 1, 42, 'text',
                    u"""Đơn vị báo cáo: %s""" % (
                        _p.return_company_info('name')),
                    None, self.rh_cell_style_header_1),
                   ('heading_2', 1, 10, 'text',
                    ' ',
                    None, self.rh_cell_style_header_1),
                   ('heading_3', 3, 60, 'text',
                    'Mẫu số B01 – DN',
                    None, self.rh_cell_style_header_2)]

        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)

        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 500
        c_specs = [
            ('heading_1', 1, 42, 'text',
             u'Địa chỉ: %s' % (_p.return_company_info('address')),
             None, self.rh_cell_style_header_1),
            ('heading_2', 1, 10, 'text',
                ' ',
                None, self.rh_cell_style_header_1),
            ('heading_3', 3, 60, 'text',
                '(Ban hành theo Thông tư số 200/2014/TT-BTC '
                'Ngày 22 / 12 /2014 của Bộ Tài chính)',
                None, self.rh_cell_style_header_3_2)
        ]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)

        c_specs = [('heading_1', 1, 42, 'text',
                    ' ',
                    None, self.rh_cell_style_header_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 2 * 256
        c_specs = [('heading_1', 5, 42, 'text',
                    'BẢNG CÂN ĐỐI KẾ TOÁN  ',
                    None, self.rh_cell_style_header_4)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_4)
        c_specs = [('heading_1', 5, 42, 'text',
                    u'Tại %s' % (_p. get_date(data)),
                    None, self.rh_cell_style_header_3_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_3)
        c_specs = [('heading_1', 1, 42, 'text',
                    _p.compute_data(data),
                    None, self.rh_cell_style_header_1_1)]
        row_pos = self.write_line(ws, row_pos, c_specs, None)
        c_specs = [('heading_1', 5, 42, 'text',
                    u'Đơn vị tính: %s' % (_p. return_company_info('currency')),
                    None, self.rh_cell_style_header_1_2)]
        row_pos = self.write_line(ws, row_pos, c_specs, None)
        for line in define_template:
            style = 'borders: ' + \
                    'left thin, right thin, top thin, bottom thin;' + \
                    _xs['wrap'] + \
                    'font: name Arial ;'
            if line[5] == '1.1':
                style += _xs['bold'] + _xs['center']
                c_specs = [
                    (x, 1, 50, 'text', x, None,
                     xlwt.easyxf(style)) for x in line[:-1]
                ]
            elif line[5] == '1.2':
                style += _xs['center']
                c_specs = [
                    (x, 1, 50, 'number', x, None,
                        xlwt.easyxf(style, num_format_str='#,##0 ;(#,##0)'))
                    for x in line[:-1]
                ]
            elif line[5] == '2.1.1':
                style += _xs['bold']
                c_specs = [
                    (line[0], 1, 50, 'text', line[0], None,
                     xlwt.easyxf(style + _xs['left'])),
                    (line[1], 1, 10, 'text', line[1], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[2], 1, 25, 'text', line[2], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[3], 1, 25, 'number', None, line[3],
                     xlwt.easyxf(style + _xs['right'],
                                 num_format_str='#,##0 ;(#,##0)')),
                    (line[4], 1, 30, 'number', None, line[4],
                     xlwt.easyxf(style + _xs['right'],
                                 num_format_str='#,##0 ;(#,##0)'))
                ]
            elif line[5] == '2.1.2':
                style += _xs['bold']
                c_specs = [
                    (line[0], 1, 50, 'text', line[0], None,
                     xlwt.easyxf(style + _xs['left'])),
                    (line[1], 1, 10, 'text', line[1], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[2], 1, 25, 'text', line[2], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[3], 1, 25, 'number', render(line[3]),
                     None, xlwt.easyxf(style + _xs['right'],
                                       num_format_str='#,##0 ;(#,##0)')),
                    (line[4], 1, 30, 'number', render(line[4]),
                     None, xlwt.easyxf(style + _xs['right'],
                                       num_format_str='#,##0 ;(#,##0)'))
                ]
            elif line[5] == '2.2':
                c_specs = [
                    (line[0], 1, 50, 'text', line[0], None,
                     xlwt.easyxf(style + _xs['left'])),
                    (line[1], 1, 10, 'text', line[1], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[2], 1, 25, 'text', line[2], None,
                     xlwt.easyxf(style + _xs['center'])),
                    (line[3], 1, 25, 'number', render(line[3]),
                     None, xlwt.easyxf(style + _xs['right'],
                                       num_format_str='#,##0 ;(#,##0)')),
                    (line[4], 1, 30, 'number', render(line[4]),
                     None, xlwt.easyxf(style + _xs['right'],
                                       num_format_str='#,##0 ;(#,##0)'))
                ]
            else:
                c_specs = [
                    (x, 1, 55, 'text', x, None,
                     xlwt.easyxf(style)) for x in line[:-1]]
            row_pos = self.write_line(ws, row_pos, c_specs, None)
        # footer
        c_specs = [('heading_1', 5, 50, 'text',
                    '',
                    None, self.rh_cell_style_header_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 1, 42, 'text',
                    '',
                    None, self.rh_cell_style_header_1),
                   ('heading_2', 1, 10, 'text',
                    ' ',
                    None, self.rh_cell_style_header_1),
                   ('heading_3', 1, 25, 'text',
                    ' ',
                    None, self.rh_cell_style_header_1),
                   ('heading_4', 2, 60, 'text',
                    'Lập, ngày ... tháng ... năm .........',
                    None, self.rh_cell_style_header_3)]

        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 1, 42, 'text',
                    'Người lập biểu',
                    None, self.rh_cell_style_header_2),
                   ('heading_2', 2, 10, 'text',
                    'Kế toán trưởng',
                    None, self.rh_cell_style_header_2),
                   ('heading_3', 2, 25, 'text',
                    'Giám đốc',
                    None, self.rh_cell_style_header_2)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 1, 50, 'text',
                    '(Ký, họ tên)',
                    None, self.rh_cell_style_header_3),
                   ('heading_2', 2, 10, 'text',
                    '(Ký, họ tên)',
                    None, self.rh_cell_style_header_3),
                   ('heading_3', 2, 25, 'text',
                    '(Ký, họ tên, đóng dấu)',
                    None, self.rh_cell_style_header_3)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 1, 50, 'text',
                    'Ghi chú:',
                    None, self.rh_cell_style_header_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 5, 50, 'text',
                    '(1) Những chỉ tiêu không có số liệu có thể không '
                    'phải trình bày nhưng không được đánh lại số thứ tự '
                    'chỉ tiêu và “Mã số“.',
                    None, self.rh_cell_style_header_1_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 5, 50, 'text',
                    '(2) Số liệu trong các chỉ tiêu có dấu (*) được ghi '
                    'bằng số âm dưới hình thức ghi trong ngoặc đơn (...). ',
                    None, self.rh_cell_style_header_1_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)
        c_specs = [('heading_1', 5, 50, 'text',
                    '(3) Đối với doanh nghiệp có kỳ kế toán năm là năm '
                    'dương lịch (X) thì “Số cuối năm“ có thể ghi '
                    'là “31.12.X“; “Số đầu năm“ có thể ghi là “01.01.X“. ',
                    None, self.rh_cell_style_header_1_1)]
        row_pos = self.write_line(
            ws, row_pos, c_specs, self.rh_cell_style_header_1)

        return row_pos

    def write_line(self, ws, row_pos, c_specs, row_style):
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=row_style,
                                     set_column_size=True)
        return row_pos


balance_sheet_report('report.balance_sheet_report',
                     'account.move',
                     parser=balance_sheet_report_parser)


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
