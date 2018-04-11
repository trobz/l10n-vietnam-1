# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import fields
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.api import Environment
from openerp import SUPERUSER_ID

FORMAT_BOLD = '00'
FORMAT_NORMAL = '01'
FORMAT_RIGHT = '02'


class AccountBalanceSheetReportXlsx(ReportXlsx):

    def create_xlsx_report(self, ids, data, report):
        self.env = Environment(self.env.cr, SUPERUSER_ID, self.env.context)
        return super(
            AccountBalanceSheetReportXlsx, self
        ).create_xlsx_report(ids, data, report)

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]

        # choose responsible user of company on wizard
#         user_id = self.object.company_id.resp_user_id.id
#         self.object = self.object.sudo(user=user_id)
#         self.env = Environment(self.env.cr, user_id, self.env.context)

        self.sheet = workbook.add_worksheet()
        self._define_formats(workbook)
        self.setup_config()

        # Generate header
        self.generate_company_info()
        self.generate_report_template()

        # Generate main content
        self.generate_report_title()
        self.generate_main_content_header()
        self.generate_assets_content()
        self.generate_liablity_content()
        self.generate_owner_equity_content()
        self.add_formular_functions()

    def setup_config(self):
        self.row_pos = 0
        self.account_data = {}
        self._set_default_format()
        self._prepare_account_data()

    def _set_default_format(self):
        # Set width default
        self.sheet.set_row(3, 20)
        self.sheet.set_column('A:Z', None, self.format_default)

        self.sheet.set_column('A:A', 40)
        self.sheet.set_column('B:B', 10)
        self.sheet.set_column('C:C', 20)
        self.sheet.set_column('D:D', 20)
        self.sheet.set_column('E:E', 30)

    def _prepare_account_data(self):
        account_ids = self.get_account_ids()
        state = ('posted', 'draft')

        if self.object.target_move == 'posted':
            state = ('posted', '1')

        params = {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'state': state,
            'account_ids': tuple(account_ids),
        }

        sql = """
            WITH account_balance_by_partner AS (
                    SELECT ml.account_id,
                        ml.partner_id,
                        COALESCE(SUM(
                            CASE WHEN ml.date < '%(date_from)s'
                                THEN balance END),0)
                            AS begin_balance,

                        SUM(COALESCE(balance, 0)) AS end_balance
                    FROM account_move_line ml
                    JOIN account_move m
                        ON ml.move_id = m.id

                    WHERE state IN %(state)s
                        AND ml.account_id IN %(account_ids)s
                        AND ml.date <= '%(date_to)s'
                GROUP BY ml.account_id, ml.partner_id
                ORDER BY ml.account_id
            )

            SELECT account_id,
                SUM(begin_balance) AS begin_balance,

                SUM(CASE WHEN begin_balance > 0
                    THEN begin_balance ELSE 0 END)
                    AS begin_debit_balance_partner,

                SUM(CASE WHEN begin_balance < 0
                    THEN begin_balance ELSE 0 END)
                    AS begin_credit_balance_partner,

                SUM(end_balance) AS end_balance,

                SUM(CASE WHEN end_balance > 0
                    THEN end_balance ELSE 0 END)
                    AS end_debit_balance_partner,

                SUM(CASE WHEN end_balance < 0
                THEN end_balance ELSE 0 END)
                AS end_credit_balance_partner

            FROM account_balance_by_partner
            GROUP BY account_id
        """ % params

        self.env.cr.execute(sql)

        for line in self.env.cr.dictfetchall():
            account_id = line['account_id']
            self.account_data.update({
                account_id: {
                    'begin_balance': line['begin_balance'],
                    'begin_debit_balance_partner':
                        line['begin_debit_balance_partner'],
                    'begin_credit_balance_partner':
                        line['begin_credit_balance_partner'],
                    'end_balance': line['end_balance'],
                    'end_debit_balance_partner':
                        line['end_debit_balance_partner'],
                    'end_credit_balance_partner':
                        line['end_credit_balance_partner'],
                }
            })

    def get_account_ids(self, account_types=[0, 1, 2, 3, 4, 8, 9]):
        # we only need accounts which start with [0xx,1xx,2xx,3xx,4xx]
        Account = self.object.env['account.account']
        account_ids = []

        company = self.object.env.user.company_id

        for acc_type in account_types:
            account_ids += Account.search([
                ('code', '=like', '%s%%' % acc_type),
                ('company_id', '=', company.id)]).ids

        return account_ids

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()

        self.sheet.write(
            'A1', u'Đơn vị báo cáo: ' + company_info['name'],
            self.format_company_info)

        self.sheet.write(
            'A2', u'Địa chỉ: ' + company_info['address'],
            self.format_company_info)

        self.sheet.write(
            'A3', u'MST: ', self.format_company_info)

    def generate_report_template(self):
        self.sheet.merge_range(
            'C1:E1', u'Mẫu số B01 - DN', self.format_template_title)

        msg = u'(Ban hành theo Thông tư số 200/2014/TT-BTC ' + \
            u'Ngày 22 / 12 /2014 của Bộ Tài chính)'
        self.sheet.merge_range(
            'C2:E2', msg, self.format_template_desc)

    def generate_report_title(self):
        self.sheet.merge_range(
            'A4:E4', u'BẢNG CÂN ĐỐI KẾ TOÁN  ', self.format_report_title)

        date_to = fields.Date.from_string(self.object.date_to)
        self.sheet.merge_range(
            'A5:E5', u'Tại ngày %(day)s tháng %(month)s năm %(year)s' %
            {'day': date_to.day, 'month': date_to.month, 'year': date_to.year},
            self.format_report_date)

        self.sheet.write('E7', u'Đơn vị tính: VND', self.format_uom)

    def generate_main_content_header(self):
        headers = [
            u'Chi tiêu',
            u'Mã số',
            u'Thuyết minh',
            u'Số cuối kỳ',
            u'Số đầu kỳ',
        ]

        sub_headers = ['1', '2', '3', '4', '5']

        col = 0
        for index in range(0, 5):
            self.sheet.write(
                7, col, headers[index], self.format_table_bold_center)

            self.sheet.write(
                8, col, sub_headers[index], self.format_table_center)

            col += 1

    def generate_assets_content(self):
        self.sheet.write(
            'A10', u'TÀI SẢN',
            self.format_table_bold_center)
        self.row_position = 11

        self.generate_section_A_I()
        self.generate_section_A_II()
        self.generate_section_A_III()
        self.generate_section_A_IV()
        self.generate_section_A_V()

        self.generate_section_B_I()
        self.generate_section_B_II()
        self.generate_section_B_III()
        self.generate_section_B_IV()
        self.generate_section_B_V()
        self.generate_section_B_VI()

    def generate_section_A_I(self):
        content = [
            (u'A - TÀI SẢN NGẮN HẠN (100=110+120+130+140+150)',
             '100', '', FORMAT_BOLD),
            (u'I. Tiền và các khoản tương đương tiền',
             '110', '', FORMAT_BOLD),
            (u' 1.Tiền ', '111', '',
             FORMAT_RIGHT, 111),
            (u' 2. Các khoản tương đương tiền ', '112', '',
             FORMAT_NORMAL, 112),
        ]

        self.write_row(content)

    def generate_section_A_II(self):
        content = [
            (u'II. Các khoản đầu tư tài chính ngắn hạn',
             '120', '', FORMAT_BOLD),
            (u' 1. Chứng khoán kinh doanh', '121', '', FORMAT_NORMAL, 121),
            (u' 2. Dự phòng giảm giá chứng khoán kinh doanh',
             '122', '', FORMAT_NORMAL, 122),
            (u' 3. Đầu tư nắm giữ đến ngày đáo hạn',
             '123', '', FORMAT_NORMAL, 123),
        ]

        self.write_row(content)

    def generate_section_A_III(self):
        content = [
            (u'III. Các khoản phải thu ngắn hạn', '130', '', FORMAT_BOLD),
            (u' 1. Phải thu ngắn hạn của khách hàng',
             '131', '', FORMAT_NORMAL, 131),
            (u' 2. Trả trước cho người bán ngắn hạn',
             '132', '', FORMAT_NORMAL, 132),
            (u' 3. Phải thu nội bộ ngắn hạn', '133', '', FORMAT_NORMAL, 133),
            (u' 4. Phải thu theo tiến độ kế hoạch hợp đồng xây dựng',
             '134', '', FORMAT_NORMAL, 134),
            (u' 5. Phải thu về cho vay ngắn hạn',
             '135', '', FORMAT_NORMAL, 135),
            (u' 6. Phải thu ngắn hạn khác', '136', '', FORMAT_NORMAL, 136),
            (u' 7. Dự phòng phải thu ngắn hạn khó đòi (*)',
             '137', '', FORMAT_NORMAL, 137),
            (u' 8. Tài sản thiếu chờ xử lý', '138', '', FORMAT_NORMAL, 138),
        ]

        self.write_row(content)

    def generate_section_A_IV(self):
        content = [
            (u'IV. Hàng tồn kho', '140', '', FORMAT_BOLD),
            (u' 1. Hàng tồn kho', '141', '', FORMAT_NORMAL, 141),
            (u' 2. Dự phòng giảm giá hàng tồn kho (*)',
             '142', '', FORMAT_NORMAL, 142),
        ]

        self.write_row(content)

    def generate_section_A_V(self):
        content = [
            (u'V. Tài sản ngắn hạn khác', '150', '', FORMAT_BOLD),
            (u' 1. Chi phí trả trước ngắn hạn',
             '151', '', FORMAT_NORMAL, 151),
            (u' 2. Thuế GTGT được khấu trừ', '152', '', FORMAT_NORMAL, 152),
            (u' 3. Thuế và các khoản khác phải thu Nhà nước',
             '153', '', FORMAT_NORMAL, 153),
            (u' 4.  Giao dịch mua bán lại trái phiếu Chính Phủ',
             '154', '', FORMAT_NORMAL, 154),
            (u' 5. Tài sản ngắn hạn khác', '155', '', FORMAT_NORMAL, 155),
        ]

        self.write_row(content)

    def generate_section_B_I(self):
        content = [

            (u'B - TÀI SẢN DÀI HẠN (200 = 210 + 220 + 240 + 250 + 260)',
             '200', '', FORMAT_BOLD),
            (u'I- Các khoản phải thu dài hạn', '210', '', FORMAT_BOLD),
            (u' 1. Phải thu dài hạn của khách hàng',
             '211', '', FORMAT_NORMAL, 211),
            (u' 2. Trả trước cho người bán dài hạn',
             '212', '', FORMAT_NORMAL, 212),
            (u' 3. Vốn kinh doanh ở đơn vị trực thuộc',
             '213', '', FORMAT_NORMAL, 213),
            (u' 4. Phải thu nội bộ dài hạn ', '214', '', FORMAT_NORMAL, 214),
            (u' 5. Phải thu về cho vay dài hạn',
             '215', '', FORMAT_NORMAL, 215),
            (u' 6. Phải thu dài hạn khác ', '216', '', FORMAT_NORMAL, 216),
            (u' 7. Dự phòng phải thu dài hạn khó đòi (*) ',
             '219', '', FORMAT_NORMAL, 219),
        ]

        self.write_row(content)

    def generate_section_B_II(self):
        content = [
            (u'II. Tài sản cố định', '220', '', FORMAT_BOLD),
            (u' 1. Tài sản cố định hữu hình', '221', ''),
            (u'  - Nguyên giá', '222', '', FORMAT_NORMAL, 222),
            (u'  - Giá trị hao mòn luỹ kế (*)',
             '223', '', FORMAT_NORMAL, 223),
            (u' 2. Tài sản cố định thuê tài chính', '224', ''),
            (u'  - Nguyên giá', '225', '', FORMAT_NORMAL, 225),
            (u'  - Giá trị hao mòn luỹ kế (*)',
             '226', '', FORMAT_NORMAL, 226),
            (u' 3. Tài sản cố định vô hình', '227', '',),
            (u'  - Nguyên giá', '228', '', FORMAT_NORMAL, 228),
            (u'  - Giá trị hao mòn luỹ kế (*)',
             '229', '', FORMAT_NORMAL, 229),
        ]

        self.write_row(content)

    def generate_section_B_III(self):
        content = [
            (u'III. Bất động sản đầu tư', '230', '', FORMAT_BOLD),
            (u'  - Nguyên giá', '231', '', FORMAT_NORMAL, 231),
            (u'  - Giá trị hao mòn luỹ kế (*)',
             '232', '', FORMAT_NORMAL, 232),
        ]

        self.write_row(content)

    def generate_section_B_IV(self):
        content = [
            (u'IV. Tài sản dở dang dài hạn', '240', '', FORMAT_BOLD),
            (u'1. Chi phí sản xuất, kinh doanh dở dang dài hạn',
             '241', '', FORMAT_NORMAL, 241),
            (u'2. Chi phí xây dựng cơ bản dở dang',
             '242', '', FORMAT_NORMAL, 242),
        ]

        self.write_row(content)

    def generate_section_B_V(self):
        content = [
            (u'V. Các khoản đầu tư tài chính dài hạn', '250', '', FORMAT_BOLD),
            (u'1. Đầu tư vào công ty con ', '251', '', FORMAT_NORMAL, 251),
            (u'2. Đầu tư vào công ty liên kết, liên doanh ',
             '252', '', FORMAT_NORMAL, 252),
            (u'3. Đầu tư góp vốn vào đơn vị khác',
             '253', '', FORMAT_NORMAL, 253),
            (u'4. Dự phòng giảm giá đầu tư tài chính dài hạn (*) ',
             '254', '', FORMAT_NORMAL, 254),
            (u'5. Đầu tư nắm giữ đến ngày đáo hạn',
             '255', '', FORMAT_NORMAL, 255),
        ]

        self.write_row(content)

    def generate_section_B_VI(self):
        content = [
            (u'VI. Tài sản dài hạn khác', '260', '', FORMAT_BOLD),
            (u'1. Chi phí trả trước dài hạn ', '261', '', FORMAT_NORMAL, 261),
            (u'2. Tài sản thuế thu nhập hoãn lại ',
             '262', '', FORMAT_NORMAL, 262),
            (u'3. Thiết bị, vật tư, phụ tùng thay thế dài hạn',
             '263', '', FORMAT_NORMAL, 263),
            (u'4. Tài sản dài hạn khác ', '268', '', FORMAT_NORMAL, 268),
            (u'TỔNG CỘNG TÀI SẢN (270 = 100 + 200)', '270', '', FORMAT_BOLD),
        ]

        self.write_row(content)

    def generate_liablity_content(self):
        self.row_position += 1

        self.sheet.write(
            'A%s' % self.row_position, u'NGUỒN VỐN',
            self.format_table_bold_center)
        self.sheet.write(
            'B%s' % self.row_position, '',
            self.format_table_bold_center)
        self.sheet.write(
            'C%s' % self.row_position, '',
            self.format_table_bold_center)
        self.sheet.write(
            'D%s' % self.row_position, '',
            self.format_table_bold_center)
        self.sheet.write(
            'E%s' % self.row_position, '',
            self.format_table_bold_center)

        self.row_position += 1

        self.generate_section_liability_A_I()
        self.generate_section_liability_A_II()

    def generate_section_liability_A_I(self):
        content = [
            (u'A. NỢ PHẢI TRẢ (300 = 310 + 330)', '300', '', FORMAT_BOLD),
            (u'I. Nợ ngắn hạn', '310', '', FORMAT_BOLD),
            (u'1. Phải trả người bán ngắn hạn',
             '311', '', FORMAT_NORMAL, 311),
            (u'2. Người mua trả tiền trước ngắn hạn',
             '312', '', FORMAT_NORMAL, 312),
            (u'3. Thuế và các khoản phải nộp nhà nước',
             '313', '', FORMAT_NORMAL, 313),
            (u'4. Phải trả người lao động', '314', '', FORMAT_NORMAL, 314),
            (u'5. Chi phí phải trả ngắn hạn', '315', '', FORMAT_NORMAL, 315),
            (u'6. Phải trả nội bộ ngắn hạn', '316', '', FORMAT_NORMAL, 316),
            (u'7. Phải trả theo tiến độ kế hoạch hợp đồng xây dựng',
             '317', '', FORMAT_NORMAL, 317),
            (u'8. Doanh thu chưa thực hiện ngắn hạn',
             '318', '', FORMAT_NORMAL, 318),
            (u'9. Phải trả ngắn hạn khác', '319', '', FORMAT_NORMAL, 319),
            (u'10. Vay và nợ thuê tài chính ngắn hạn',
             '320', '', FORMAT_NORMAL, 320),
            (u'11. Dự phòng phải trả ngắn hạn',
             '321', '', FORMAT_NORMAL, 321),
            (u'12. Quỹ khen thưởng, phúc lợi', '322', '', FORMAT_NORMAL, 322),
            (u'13. Quỹ bình ổn giá', '323', '', FORMAT_NORMAL, 323),
            (u'14. Giao dịch mua bán lại trái phiếu Chính Phủ',
             '324', '', FORMAT_NORMAL, 324),
        ]

        self.write_row(content)

    def generate_section_liability_A_II(self):
        content = [
            (u'II. Nợ dài hạn', '330', '', FORMAT_BOLD),
            (u' 1. Phải trả dài hạn người bán',
             '331', '', FORMAT_NORMAL, 331),
            (u' 2. Người mua trả tiền trước dài hạn',
             '332', '', FORMAT_NORMAL, 332),
            (u' 3. Chi phí phải trả dài hạn', '333', '', FORMAT_NORMAL, 333),
            (u' 4. Phải trả nội bộ về vốn kinh doanh',
             '334', '', FORMAT_NORMAL, 334),
            (u' 5. Phải trả nội bộ dài hạn', '335', '', FORMAT_NORMAL, 335),
            (u' 6. Doanh thu chưa thực hiện dài hạn',
             '336', '', FORMAT_NORMAL, 336),
            (u' 7. Phải trả dài hạn khác', '337', '', FORMAT_NORMAL, 337),
            (u' 8. Vay và nợ thuê tài chính dài hạn',
             '338', '', FORMAT_NORMAL, 338),
            (u' 9. Trái phiếu chuyển đổi', '339', '', FORMAT_NORMAL, 339),
            (u' 10. Cổ phiếu ưu đãi', '340', '', FORMAT_NORMAL, 340),
            (u' 11.  Thuế thu nhập hoãn lại phải trả',
             '341', '', FORMAT_NORMAL, 341),
            (u' 12. Dự phòng phải trả dài hạn',
             '342', '', FORMAT_NORMAL, 342),
            (u' 13. Quỹ phát triển khoa học và công nghệ',
             '343', '', FORMAT_NORMAL, 343),
        ]

        self.write_row(content)

    def generate_owner_equity_content(self):
        self.generate_equity_B_I()
        self.generate_equity_B_II()

    def generate_equity_B_I(self):
        content = [
            (u'B - VỐN CHỦ SỞ HỮU (400 = 410 + 430)', '400', '', FORMAT_BOLD),
            (u'I. Vốn chủ sở hữu', '410', '', FORMAT_BOLD),
            (u' 1. Vốn chủ sở hữu', '411', '', FORMAT_NORMAL),
            (u'   - Cổ phiếu phổ thông có quyền biểu quyết',
             '411a', '', FORMAT_NORMAL, '411a'),
            (u'   - Cổ phiếu ưu đãi', '411b', '', FORMAT_NORMAL, '411b'),
            (u' 2. Thặng dư vốn cổ phần', '412', '', FORMAT_NORMAL, 412),
            (u' 3. Quyền chọn chuyển đổi trái phiếu',
             '413', '', FORMAT_NORMAL, 413),
            (u' 4. Vốn khác của chủ sở hữu ', '414', '', FORMAT_NORMAL, 414),
            (u' 5. Cổ phiếu quỹ (*) ', '415', '', FORMAT_NORMAL, 415),
            (u' 6. Chênh lệch đánh giá lại tài sản',
             '416', '', FORMAT_NORMAL, 416),
            (u' 7. Chênh lệch tỷ giá hối đoái',
             '417', '', FORMAT_NORMAL, 417),
            (u' 8. Quỹ đầu tư phát triển', '418', '', FORMAT_NORMAL, 418),
            (u' 9. Quỹ hỗ trợ sắp xếp doanh nghiệp',
             '419', '', FORMAT_NORMAL, 419),
            (u' 10. Quỹ khác thuộc vốn chủ sở hữu',
             '420', '', FORMAT_NORMAL, 420),
            (u' 11. Lợi nhuận sau thuế chưa phân phối',
             '421', '', FORMAT_NORMAL),
            (u'    - LNST chưa phân phối lũy kế đến cuối kỳ trước',
             '421a', '', FORMAT_NORMAL, '421a'),
            (u'    - LNST chưa phân phối kỳ này',
             '421b', '', FORMAT_NORMAL, '421b'),
            (u' 12. Nguồn vốn đầu tư XDCB', '422', '', FORMAT_NORMAL, 422)
        ]

        self.write_row(content)

    def generate_equity_B_II(self):
        content = [
            (u'II. Nguồn kinh phí và quỹ khác', '430', '', FORMAT_BOLD),
            (u' 1. Nguồn kinh phí', '431', '', FORMAT_NORMAL, 431),
            (u' 2. Nguồn kinh phí đã hình thành TSCĐ',
             '432', '', FORMAT_NORMAL, 432),
            (u'TỔNG CỘNG NGUỒN VỐN (440 = 300 + 400)',
             '440', '', FORMAT_BOLD, 433),
        ]

        self.write_row(content)

    def write_row(self, content):
        for item in content:
            format_cells = self.get_format(item)
            account_balance = self.get_account_balance(item)

            criteria = item[0]
            code = item[1]
            description = item[2]

            self.sheet.write(
                'A%s' % self.row_position, criteria, format_cells['left'])

            self.sheet.write(
                'B%s' % self.row_position, code, format_cells['center'])

            self.sheet.write(
                'C%s' % self.row_position, description, format_cells['center'])

            self.sheet.write(
                'D%s' % self.row_position,
                account_balance['end_balance'], format_cells['right'])

            self.sheet.write(
                'E%s' % self.row_position,
                account_balance['begin_balance'], format_cells['right'])

            self.row_position += 1

    def get_format(self, item):
        code = self.get_format_code(item)
        if code == FORMAT_BOLD:
            return {
                'left': self.format_table_bold,
                'center': self.format_table_bold_center,
                'right': self.format_table_bold_right,
            }

        return {
            'left': self.format_table,
            'center': self.format_table_center,
            'right': self.format_table_right,
        }

    def get_format_code(self, item):
        format_index = 3
        return item[format_index] if len(item) > format_index else ''

    def get_account_balance(self, item):
        balance_index = 4
        key = item[balance_index] if len(item) > balance_index else False

        return self.compute_balance(key)

    def compute_balance(self, key):
        self.account_balance = {
            'end_balance': 0.0,
            'begin_balance': 0.0,
        }

        balance_sheet_conf = self.get_balance_sheet_conf(key)
        for config_line in balance_sheet_conf.config_line_ids:
            self.config_line = config_line
            account_code = config_line.code
            account_ids = self.get_account_ids([account_code])
            accounts = self.env['account.account'].browse(account_ids)

            for account in accounts:
                self.compute_account_balance(account)

        balance = self.account_balance

        # recompute balance for config
        if balance_sheet_conf.is_inverted_result:
            for key, value in balance.iteritems():
                balance[key] = -value

        if balance_sheet_conf.is_parenthesis:
            for key, value in balance.iteritems():
                balance[key] = -abs(value)

        return balance

    def get_balance_sheet_conf(self, key):
        BalanceSheetConfig = self.env['balance.sheet.config']
        return BalanceSheetConfig.search(
            [('code', '=', key)], limit=1)

    def compute_account_balance(self, account):
        account_balance_data = self.get_formated_account_balance(account)
        account_balance_data = self.compute_for_parenthesis_balance(
            account_balance_data)
        self.compute_for_dedit_balance(account_balance_data)
        self.compute_for_credit_balance(account_balance_data)
        self.compute_for_inverted_balance(account_balance_data)

    def get_formated_account_balance(self, account):
        account_id = account.id
        account_dict = self.account_data.get(account_id)

        result = {
            'begin_debit_balance': 0,
            'begin_credit_balance': 0,
            'end_debit_balance': 0,
            'end_credit_balance': 0,
        }

        if not account_dict:
            return result

        if account.internal_type in ('receivable', 'payable'):
            result['begin_debit_balance'] = \
                account_dict['begin_debit_balance_partner']
            result['begin_credit_balance'] = \
                account_dict['begin_credit_balance_partner']
            result['end_debit_balance'] = \
                account_dict['end_debit_balance_partner']
            result['end_credit_balance'] = \
                account_dict['end_credit_balance_partner']

        else:
            begin_balance = account_dict['begin_balance']
            if begin_balance > 0:
                result['begin_debit_balance'] = begin_balance
                result['begin_credit_balance'] = 0
            else:
                result['begin_credit_balance'] = 0
                result['begin_credit_balance'] = abs(begin_balance)

            end_balance = account_dict['end_balance']
            if end_balance > 0:
                result['end_debit_balance'] = end_balance
                result['end_credit_balance'] = 0
            else:
                result['end_debit_balance'] = 0
                result['end_credit_balance'] = abs(end_balance)

        return result

    def compute_for_parenthesis_balance(self, account_balance_data):
        if self.config_line.is_parenthesis:
            for key, value in account_balance_data.iteritems():
                account_balance_data[key] = -value

        return account_balance_data

    def compute_for_dedit_balance(self, account_balance_data):
        config_line = self.config_line
        if config_line.is_debit_balance and \
                not config_line.is_credit_balance:

            if config_line.operator == 'plus':
                self.account_balance['begin_balance'] += \
                    account_balance_data['begin_debit_balance']

                self.account_balance['end_balance'] += \
                    account_balance_data['end_debit_balance']

            else:
                self.account_balance['begin_balance'] -= \
                    account_balance_data['begin_debit_balance']

                self.account_balance['end_balance'] -= \
                    account_balance_data['end_debit_balance']

    def compute_for_credit_balance(self, account_balance_data):
        config_line = self.config_line
        if not config_line.is_debit_balance and \
                config_line.is_credit_balance:

            if config_line.operator == 'plus':
                self.account_balance['begin_balance'] += \
                    account_balance_data['begin_credit_balance']

                self.account_balance['end_balance'] += \
                    account_balance_data['end_credit_balance']

            else:
                self.account_balance['begin_balance'] -= \
                    account_balance_data['begin_credit_balance']

                self.account_balance['end_balance'] -= \
                    account_balance_data['end_credit_balance']

    def compute_for_inverted_balance(self, account_balance_data):
        """
        If account is not inverted and has debit balance,
        the balance will be positive and vice versa.
        """
        config_line = self.config_line
        if config_line.is_debit_balance and \
                config_line.is_credit_balance:

            index = 1
            if config_line.is_inverted:
                index = -1

            if config_line.operator == 'plus':
                begin_debit_balance = \
                    account_balance_data['begin_debit_balance']

                if begin_debit_balance:
                    self.account_balance['begin_balance'] += \
                        account_balance_data['begin_debit_balance'] * index
                else:
                    self.account_balance['begin_balance'] += \
                        account_balance_data['begin_credit_balance'] * (-index)

                end_debit_balance = \
                    account_balance_data['end_debit_balance']

                if end_debit_balance:
                    self.account_balance['end_balance'] += \
                        account_balance_data['end_debit_balance'] * index
                else:
                    self.account_balance['end_balance'] += \
                        account_balance_data['end_credit_balance'] * (-index)

            else:
                begin_debit_balance = \
                    account_balance_data['begin_debit_balance']

                if begin_debit_balance:
                    self.account_balance['begin_balance'] -= \
                        account_balance_data['begin_debit_balance'] * index
                else:
                    self.account_balance['begin_balance'] -= \
                        account_balance_data['begin_credit_balance'] * (-index)

                end_debit_balance = \
                    account_balance_data['end_debit_balance']

                if end_debit_balance:
                    self.account_balance['end_balance'] -= \
                        account_balance_data['end_debit_balance'] * index
                else:
                    self.account_balance['end_balance'] -= \
                        account_balance_data['end_credit_balance'] * (-index)

    def add_formular_functions(self):
        # ---------------------------------------------------------------------
        # ASSETS
        # ---------------------------------------------------------------------
        # Section A:
        self.sheet.write_formula(
            'D11', '= SUM(D12, D15, D19, D28, D31)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'D12', '= SUM(D13:D14)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D15', '= SUM(D16:D18)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D19', '= SUM(D20:D27)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D28', '= SUM(D29:D30)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D28', '= SUM(D29:D30)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D31', '= SUM(D32:D36)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E11', '= SUM(E12, E15, E19, E28, E31)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'E12', '= SUM(E13:E14)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E15', '= SUM(E16:E18)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E19', '= SUM(E20:E27)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E28', '= SUM(E29:E30)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E28', '= SUM(E29:E30)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E31', '= SUM(E32:E36)', self.format_table_bold_right)

        # Section B
        self.sheet.write_formula(
            'D37', '= SUM(D38,D46, D56, D59, D62, D68)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'D38', '= SUM(D39:D45)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D46', '= SUM(D47, D50, D53)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D47', '= SUM(D48:D49)', self.format_table_right)

        self.sheet.write_formula(
            'D50', '= SUM(D51:D52)', self.format_table_right)

        self.sheet.write_formula(
            'D53', '= SUM(D54:D55)', self.format_table_right)

        self.sheet.write_formula(
            'D56', '= SUM(D57:D58)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D59', '= SUM(D60:D61)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D62', '= SUM(D63:D67)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D68', '= SUM(D69:D72)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D73', '= SUM(D11, D37)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E37', '= SUM(E38,E46, E56, E59, E62, E68)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'E38', '= SUM(E39:E45)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E46', '= SUM(E47, E50, E53)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E47', '= SUM(E48:E49)', self.format_table_right)

        self.sheet.write_formula(
            'E50', '= SUM(E51:E52)', self.format_table_right)

        self.sheet.write_formula(
            'E53', '= SUM(E54:E55)', self.format_table_right)

        self.sheet.write_formula(
            'E56', '= SUM(E57:E58)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E59', '= SUM(E60:E61)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E62', '= SUM(E63:E67)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E68', '= SUM(E69:E72)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E73', '= SUM(E11, E37)', self.format_table_bold_right)

        # ---------------------------------------------------------------------
        # LIABILITY
        # ---------------------------------------------------------------------
        self.sheet.write_formula(
            'D76', '= SUM(D77, D92)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D77', '= SUM(D78:D91)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D92', '= SUM(D93:D105)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E76', '= SUM(E77, E92)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E77', '= SUM(E78:E91)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E92', '= SUM(E93:E105)', self.format_table_bold_right)

        # ---------------------------------------------------------------------
        # EQUITY
        # ---------------------------------------------------------------------
        self.sheet.write_formula(
            'D106', '= SUM(D107,D123)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D106', '= SUM(D107,D124)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D107', '= SUM(D108, D111:D120, D123)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'D108', '= SUM(D109:D110)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D120', '= SUM(D121:D122)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D124', '= SUM(D125:D126)', self.format_table_bold_right)

        self.sheet.write_formula(
            'D127', '= SUM(D76,D106)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E106', '= SUM(E107,E123)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E106', '= SUM(E107,E124)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E107', '= SUM(E108, E111:E120, E123)',
            self.format_table_bold_right)

        self.sheet.write_formula(
            'E108', '= SUM(E109:E110)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E120', '= SUM(E121:E122)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E124', '= SUM(E125:E126)', self.format_table_bold_right)

        self.sheet.write_formula(
            'E127', '= SUM(E76,E106)', self.format_table_bold_right)

    def _define_formats(self, workbook):
        """
        Add cell formats to current workbook.
        Those formats can be used on all cell.
        """
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        format_config = {
            'font_name': 'Arial',
            'font_size': 10,
        }
        self.format_default = workbook.add_format(format_config)

        format_bold = format_config.copy()
        format_bold.update({
            'bold': True,
        })
        self.format_bold = workbook.add_format(format_bold)

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
            'num_format': '#,##0;(#,##0)',
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


AccountBalanceSheetReportXlsx(
    'report.balance_sheet_report', 'account.balance.sheet')
