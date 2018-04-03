# -*- coding: utf-8 -*-
from . import general_ledger_report
import xlwt
from openerp.addons.report_xls.utils import _render # @UnresolvedImport
from .. import report_xls_utils

class Parser(general_ledger_report.Parser):
    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)

        self.report_name = 'account_ledger_report_xls'
        self.localcontext.update({

        })

class account_ledger_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(account_ledger_xls, self).__init__(name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A','B','C','D','E', 'F', 'G', 'H', 'I']
        self.col_specs_template = {
            'A': {
                'lines': [1, 12, _render("line.get('create_date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('create_date','')[:10],'%Y-%m-%d')"), None, self.style_date_right],
                },
            'B': {
                'lines': [1, 17, 'text', _render("line.get('ref','') or ''"), None, self.normal_style_left_borderall],
                },
            'C': {
                'lines': [1, 17, _render("line.get('date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('date','')[:10],'%Y-%m-%d')"), None, self.style_date_right],
                },
            'D': {
                'lines': [1, 50, 'text', _render("line.get('description','') or ''"), None, self.normal_style_left_borderall],
                },
            'E': {
                'lines': [1, 15, 'text', _render("line.get('counterpart_account',None)"), None, self.normal_style_left_borderall],
                },
            'F': {
                'lines': [1, 28, _render("line.get('debit') and 'number' or 'text'"), _render("line.get('debit','')"), None, self.style_decimal],
                },
            'G': {
                'lines': [1, 28, _render("line.get('credit') and 'number' or 'text'"), _render("line.get('credit','')"), None, self.style_decimal],
                },
            'H': {
                  'lines': [1, 28, _render("remain_debit and 'number' or 'text'"),_render("remain_debit"),  None, self.style_decimal],
                },
            'I': {
                  'lines': [1, 28, _render("remain_credit and 'number' or 'text'"), _render("remain_credit"), None,  self.style_decimal],
                }
            }

    def generate_xls_header(self, _p, _xs, data, objects, wb, ws, row_pos, report_name):
        """
        @return: row_pos: position of at the end of generatioon header
        """

        cell_address_style = self.get_cell_style(['bold', 'wrap', 'left', 'top'])
        # Title address 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 500
        c_specs = [
            ('company_name', 3, 0, 'text', u'Đơn vị: %s' % _p.get_company()['name'] or ''  , '', cell_address_style),
            ('form_serial', 4, 0, 'text', u'''Mẫu số S38 – DN'''  , '', self.normal_style_bold)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 800
        c_specs = [
            ('company_name', 3, 0, 'text', u'Địa chỉ: %s' % _p.get_company()['address'] or '', '', cell_address_style),
            ('form_serial', 4, 0, 'text', u'''(Ban hành theo QĐ số 15/2006/QĐ-BTC, Ngày 20/03/2006 của Bộ trưởng BTC)''', '', self.normal_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('company_name', 7, 0, 'text', u'MST: %s' % _p.get_company()['vat'] or '', '', cell_address_style),
            ('empty1', 1, 0, 'text', ''),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 17, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 50, 'text', '', None),
            ('col4', 1, 15, 'text', '', None),
            ('col6', 1, 28, 'text', '', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "SỔ CÁI TÀI KHOẢN"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_title_style = self.get_cell_style(['bold', 'wrap', 'center', 'middle', 'fontsize_350'])

        c_specs = [
            ('payment_journal', 7, 0, 'text', report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Tên tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        title = u'Tên tài khoản:'
        c_specs = [
            ('amount_on_account', 7, 0, 'text',u'%s %s' % (title,_p.get_account_info().get('name',''),))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)


        # Title "Sô hiệu tài khoản"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('from_to', 7, 0, 'text', u'Số hiệu tài khoản: %s' % (_p.get_account_info().get('code','')), None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)


        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('asset_book', 7, 0, 'text','Từ %s đến %s' % (_p.get_date().get('date_from','.......'),_p.get_date().get('date_to','.......')))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 10, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)


        # Header Title 1
        row_title_body_pos = row_pos
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', 'Ngày tháng ghi sổ', None),
            ('col2', 2, 34, 'text', 'Chứng Từ', None ),
            ('col3', 1, 50, 'text', 'Diễn giải', None),
            ('col4', 1, 15, 'text', 'TK đối ứng', None),
            ('col5', 2, 34, 'text', 'Số phát sinh', None),
            ('col6', 2, 34, 'text', 'Số dư', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)


        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 25, 'text', 'Số hiệu', None),
            ('col3', 1, 17, 'text', 'Ngày tháng', None),
            ('col4', 1, 50, 'text', '', None),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 17, 'text', 'Nợ', None),
            ('col7', 1, 17, 'text', 'Có', None),
            ('col8', 1, 17, 'text', 'Nợ', None),
            ('col9', 1, 17, 'text', 'Có', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 0, 0, 'Ngày tháng ghi sổ', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 3, 3, 'Diễn giải', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 4, 4, 'TK đối ứng', self.normal_style_bold_borderall )

        return row_pos

    def generate_worksheet(self, _p, _xs, data, objects, wb, report_name, count = 0):
        """
        @summary: get new worksheet from workbook, reset current row position in the new worksheet
        @return: new worksheet, new row_pos
        """
        report_name = count and (report_name[:31] + ' ' + str(count)) or report_name[:31]
        ws = wb.add_sheet(report_name, cell_overwrite_ok=True)
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_num_pages = 1
        ws.fit_height_to_pages = 0
        ws.fit_width_to_pages = 1 # allow to print fit one page
        row_pos = 0

        return ws,row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        MAX_ROW = 65500
        count = 1
        report_name = u"SỔ CHI TIẾT TÀI KHOẢN"

        # set print sheet
        ws = super(account_ledger_xls, self).generate_xls_report(_p, _xs, data, objects, wb, report_name)
        row_pos = 0

        row_pos = self.generate_xls_header(_p, _xs, data, objects, wb, ws, row_pos, report_name)

        lines_data = _p.get_lines()

        # Header số dư đầu kỳ
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        beginning_balance = _p.get_init()
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', 'Số dư đầu kỳ', None),
            ('col5', 1, 34, 'text', '', None),
            ('col6', 1, 34, 'text', '', None),
            ('col7', 1, 34, 'text', '', None),
            ('col8', 1, 50, beginning_balance.get('debit') and 'number' or 'text', beginning_balance.get('debit', ''), None, self.style_decimal_bold),
            ('col9', 1, 15, beginning_balance.get('credit') and 'number' or 'text', beginning_balance.get('credit', ''), None, self.style_decimal_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)

        sum_debit_balance = sum_credit_balance =  0

        for line in lines_data: # @UnusedVariable

            if row_pos > MAX_ROW:
                ws.flush_row_data()
                ws, row_pos = self.generate_worksheet(_p, _xs, data, objects, wb, report_name, count)
                row_pos = self.generate_xls_header(_p, _xs, data, objects, wb, ws, row_pos, report_name)
                count += 1

            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 450


            sum_debit_balance += line.get('debit', 0) or 0
            sum_credit_balance += line.get('credit', 0) or 0

            remain_debit = sum_debit_balance + beginning_balance.get('debit', 0) # @UnusedVariable
            remain_credit = sum_credit_balance + beginning_balance.get('credit', 0) # @UnusedVariable

            c_specs = map(lambda x: self.render(x, self.col_specs_template, 'lines'), self.wanted_list)
            row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
            row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_borderall)



        # Header cộng phát sinh
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', 'Cộng phát sinh', None),
            ('col5', 1, 0, 'text', '', None ),
            ('col6', 1, 0, sum_debit_balance > 0 and 'number' or 'text', sum_debit_balance or '', None, self.style_decimal_bold ),
            ('col7', 1, 0, sum_credit_balance > 0 and 'number' or 'text', sum_credit_balance or '', None, self.style_decimal_bold),
            ('col8', 1, 28, 'text', '', None),
            ('col9', 1, 28, 'text', '', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)

        # Header số dư cuối kỳ
        total_balance = (beginning_balance.get('debit', 0) + sum_debit_balance) - (beginning_balance.get('credit', 0) + sum_credit_balance)
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', 'Số dư cuối kỳ', None),
            ('col5', 1, 0, 'text', '', None ),
            ('col6', 1, 0, 'text', '', None ),
            ('col7', 1, 0, 'text', '', None ),
            ('col8', 1, 0, total_balance > 0 and 'number' or 'text', total_balance > 0 and total_balance or '' , None, self.style_decimal_bold),
            ('col9', 1, 0, total_balance < 0 and 'number' or 'text', total_balance < 0 and abs(total_balance) or '' , None, self.style_decimal_bold),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)
#
        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(5)]
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
            ('col6', 1, 0, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col9', 2, 0, 'text', 'Giám đốc', None),
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
            ('col5', 1, 0, 'text', '', None),
            ('col6', 1, 0, 'text', '(Ký, họ tên)', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=cell_footer_style)

account_ledger_xls('report.account_ledger_report_xls','account.move.line', parser=Parser)

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
