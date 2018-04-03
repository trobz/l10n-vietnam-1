# -*- coding: utf-8 -*-
import xlwt
from openerp.addons.report_base_vn.report import report_base_vn
from datetime import date, datetime
import time
from openerp.addons.report_xls.utils import _render # @UnresolvedImport
from .. import report_xls_utils
import logging
_logger = logging.getLogger(__name__)

class Parser(report_base_vn.Parser):
    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)

        self.report_name = 'cash_book_report'
        self.localcontext.update({
            'get_account': self.get_account,
            'get_wizard_data': self.get_wizard_data,
            'get_lines': self.get_lines,
            'get_init': self.get_init,
            'get_name_report': self.get_name_report,
            'get_template': self.get_template,
            'datetime': datetime
        })

    def get_wizard_data(self):
        result = {}
        datas = self.localcontext['data']
        if datas:
            result['fiscalyear'] = 'fiscalyear_id' in datas['form'] and datas['form']['fiscalyear_id'] or False
            result['chart_account_id'] = 'chart_account_id' in datas['form'] and datas['form']['chart_account_id'] or False
            result['target_move'] = 'target_move' in datas['form'] and datas['form']['target_move'] or ''
            result['type_report'] = 'type_report' in datas['form'] and datas['form']['type_report'] or False
            result['account'] = 'account' in datas['form'] and datas['form']['account'] or False
            result['filter'] = 'filter' in datas['form'] and datas['form']['filter'] or False
            if datas['form']['filter'] == 'filter_date':
                result['date_from'] = datas['form']['date_from']
                result['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                result['period_from'] = datas['form']['period_from']
                result['period_to'] = datas['form']['period_to']
            result['target_move'] = 'target_move' in datas['form'] and datas['form']['target_move'] or False

        return result

    def get_account(self):
        wizard_data = self.get_wizard_data()
        if wizard_data:
            return wizard_data['account'] and wizard_data['account'][1] or False
        return False

    def get_name_report(self):
        wizard_data = self.get_wizard_data()
        name = False
        if wizard_data.get('type_report', False):
            if wizard_data['type_report'] == 'cash_book':
                name = u'SỔ QUỸ TIỀN MẶT'
            else:
                name = u'SỔ QUỸ NGÂN HÀNG'
        return name

    def get_template(self):
        wizard_data = self.get_wizard_data()
        res = {'code': False,
               'decision_code': False,
               'date': False}

        if wizard_data.get('type_report', False):
            if wizard_data['type_report'] == 'cash_book':
                res = {'code': 'S07',
                       'decision_code': '15/2006',
                       'date': '20/03/2006'
                       }
            else:
                res = {'code': 'S08',
                       'decision_code': '15/2006',
                       'date': '20/03/2006'
                       }
        return res

    def get_init(self):
        wizard_data = self.get_wizard_data()
        date_info = self.get_date()
        acc_move_line_obj = self.pool.get('account.move.line')
        init = 0.0
        if date_info:
            if date_info['date_from'] == date(int(wizard_data['fiscalyear'][1]), 1, 1):
                return init
            if wizard_data['target_move'] == 'posted':
                state = ['posted']
            else:
                state = ['posted', 'draft' ]
            acc_move_line_ids = acc_move_line_obj.search(self.cr, self.uid, [('account_id', '=', wizard_data['account'][0]),
                                                                         ('move_id.state','in', state),
                                                                         ('date', '>=', date(int(wizard_data['fiscalyear'][1]), 1, 1)),
                                                                         ('date', '<', date_info['date_from_date']),

                                                                    ])
            account_move_line_data = acc_move_line_obj.read(self.cr, self.uid, acc_move_line_ids, ['credit', 'debit'])
            for acc_move_line in account_move_line_data:
                init = init + acc_move_line['debit'] - acc_move_line['credit']
            return init
        else:
            return init



    def get_lines(self):
        start_time = time.time()
        wizard_data = self.get_wizard_data()
        date_info = self.get_date()

        account_id = wizard_data['account'][0] or -1
        params  = { 'target_move': '', 'account_id': account_id,
                   'date_from': date_info['date_from_date'],  'date_to': date_info['date_to_date'] }

        if wizard_data['target_move'] == 'posted':
            params['target_move'] = "AND amv.state = 'posted'"


        SQL = """
            SELECT

                aml_dr.date_created AS create_date,
                aml_dr.date AS date,
                CASE WHEN  aml_dr.debit > 0   then amv.name else '' END as ref_debit,
                CASE WHEN  aml_dr.credit > 0  then amv.name else '' END as ref_credit,

                CASE WHEN  acc.id is null then acc_dr.code else acc_cr.code END as counterpart_account,

                CASE
                    WHEN acc.id is null
                        THEN aml_cr.name
                        ELSE aml_dr.name END AS description,
                amv.narration AS narration,
                CASE
                    WHEN acc.id is not null  AND aml_dr.debit > 0 THEN aml_dr.debit
                    WHEN acc.id is null      AND aml_cr.debit > 0 THEN aml_dr.credit
                END debit,

                CASE
                    WHEN acc.id is not null  AND aml_dr.credit > 0 THEN aml_dr.credit
                    WHEN acc.id is null     AND aml_cr.credit > 0 THEN aml_dr.debit
                END credit


            FROM account_move_line aml_dr

            JOIN
                account_move_line aml_cr ON aml_dr.counter_move_id = aml_cr.id

            JOIN
                account_move amv ON amv.id = aml_dr.move_id

            LEFT JOIN
                (   SELECT id
                    FROM account_account
                    where id = %(account_id)s) acc ON acc.id = aml_dr.account_id
            JOIN (SELECT id,code from account_account) acc_dr ON acc_dr.id = aml_dr.account_id
            JOIN (SELECT id,code from account_account) acc_cr ON acc_cr.id = aml_cr.account_id


            WHERE (aml_dr.account_id=%(account_id)s or aml_cr.account_id=%(account_id)s )
            and aml_dr.date >= '%(date_from)s' and aml_dr.date <= '%(date_to)s'
            %(target_move)s
            ORDER BY aml_dr.date, amv.name

            """

        SQL = SQL % params
        self.cr.execute(SQL)
        res = self.cr.dictfetchall()

        _logger.info("End process elapsed time: %s" % ( time.time() - start_time)) # debug mode
        return res

class cash_book_report_xls(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(cash_book_report_xls, self).__init__(name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A','B','C','D','E','F','G','H','I','K']
        self.col_specs_template = {
            'A': {
                'lines': [1, 12, _render("line.get('create_date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('create_date','')[:10],'%Y-%m-%d')"), None, self.style_date_right],
                },
            'B': {
                'lines': [1, 12, _render("line.get('date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('date','')[:10],'%Y-%m-%d')"), None, self.style_date_right],
                },
            'C': {
                'lines': [1, 17, 'text', _render("line.get('ref_debit','') or ''"), None, self.normal_style_left_borderall],
                },
            'D': {
                'lines': [1, 17, 'text', _render("line.get('ref_credit','') or ''"), None, self.normal_style_left_borderall],
                },
            'E': {
                'lines': [1, 50, 'text', _render("line.get('description','') or ''"), None, self.normal_style_left_borderall],
                },
            'F': {
                'lines': [1, 15, 'text', _render("line.get('counterpart_account',None)"), None, self.normal_style_left_borderall],
                },
            'G': {
                'lines': [1, 28, _render("line.get('debit') and 'number' or 'text'"), _render("line.get('debit','')"), None, self.style_decimal],
                },
            'H': {
                'lines': [1, 28, _render("line.get('credit') and 'number' or 'text'"), _render("line.get('credit','')"), None, self.style_decimal],
                },
            'I': {
                'lines': [1, 28, _render("line.get('remain') and 'number' or 'text'"), _render("line.get('remain','')"), None, self.style_decimal],
                },
            'K': {
                'lines': [1, 50, 'text', _render("line.get('narration','') or ''"), None, self.normal_style_left_borderall],
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
            ('company_name', 4, 0, 'text', u'Đơn vị: %s' % _p.get_company()['name'] or ''  , '', cell_address_style),
            ('form_serial', 6, 0, 'text', u'''Mẫu số %s – DN''' % _p.get_template()['code']  , '', self.normal_style_bold)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 800
        c_specs = [
            ('company_name', 4, 0, 'text', u'Địa chỉ: %s' % _p.get_company()['address'] or '', '', cell_address_style),
            ('form_serial', 6, 0, 'text', u'''(Ban hành theo QĐ số %s/QĐ-BTC, Ngày %s của Bộ trưởng BTC)''' % (_p.get_template()['decision_code'], _p.get_template()['date']), '', self.normal_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('company_name', 10, 0, 'text', u'MST: %s' % _p.get_company()['vat'] or '', '', cell_address_style),
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
            ('payment_journal', 10, 0, 'text', report_name, None, cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Loại quỹ"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        title = u'Loại quỹ:'
        c_specs = [
            ('amount_on_account', 10, 0, 'text',u'%s %s' % (title, _p.get_account()))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('asset_book', 10, 0, 'text','Từ %s đến %s' % (_p.get_date().get('date_from','.......'),_p.get_date().get('date_to','.......')))
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
            ('col2', 1, 12, 'text', 'Ngày tháng chứng từ', None ),
            ('col3', 2, 28, 'text', 'Chứng từ', None),
            ('col4', 1, 35, 'text', 'Diễn giải', None),
            ('col5', 1, 10, 'text', 'TK đối ứng', None),
            ('col6', 3, 34, 'text', 'Số Tiền', None),
            ('col7', 1, 34, 'text', 'Ghi chú', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)


        # Header Title 2
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 15, 'text', 'Thu', None),
            ('col4', 1, 15, 'text', 'Chi', None),
            ('col5', 1, 35, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 17, 'text', 'Thu', None),
            ('col8', 1, 17, 'text', 'Chi', None),
            ('col9', 1, 17, 'text', 'Tồn', None),
            ('col10', 1, 17, 'text', '', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 0, 0, 'Ngày tháng ghi sổ', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 1, 1, 'Ngày tháng chứng từ', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 4, 4, 'Diễn giải', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 5, 5, 'TK đối ứng', self.normal_style_bold_borderall )
        ws.write_merge(row_title_body_pos, row_title_body_pos+1, 9, 9, 'Ghi chú', self.normal_style_bold_borderall )

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
        report_name = _p.get_name_report()

        # call parent init utils.
        # set print sheet
        ws = super(cash_book_report_xls, self).generate_xls_report(_p, _xs, data, objects, wb, report_name)
        row_pos = 0

        row_pos = self.generate_xls_header(_p, _xs, data, objects, wb, ws, row_pos, report_name)

        lines_data = _p.get_lines()
        # Header số dư đầu kỳ
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        beginning_balance = _p.get_init()
        c_specs = [
            ('col1', 1, 10, 'text', '', None),
            ('col2', 1, 10, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', '', None),
            ('col5', 1, 12, 'text', 'Số dư đầu kỳ', None),
            ('col6', 1, 34, 'text', '', None ),
            ('col7', 1, 34, 'text', '', None ),
            ('col8', 1, 34, 'text', '', None ),
            ('col9', 1, 50, 'number', beginning_balance, None, self.style_decimal_bold),
            ('col10', 1, 34, 'text', '', None )
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)

        remain_balance = beginning_balance
        sum_debit_balance = sum_credit_balance = 0
        for line in lines_data: # @UnusedVariable
            if row_pos > MAX_ROW:
                ws.flush_row_data()
                ws, row_pos = self.generate_worksheet(_p, _xs, data, objects, wb, report_name, count)
                row_pos = self.generate_xls_header(_p, _xs, data, objects, wb, ws, row_pos, report_name)
                count += 1
            remain_balance = int(line.get('debit') or 0) - int(line.get('credit') or 0) + remain_balance
            sum_debit_balance += int(line.get('debit') or 0)
            sum_credit_balance += int(line.get('credit') or 0)
            line['remain'] = remain_balance
            ws.row(row_pos).height_mismatch = True
            ws.row(row_pos).height = 450
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
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', 'Cộng phát sinh', None),
            ('col6', 1, 0, 'text', '', None ),
            ('col7', 1, 0, sum_debit_balance > 0 and 'number' or 'text', sum_debit_balance or '', None, self.style_decimal_bold ),
            ('col8', 1, 0, sum_credit_balance > 0 and 'number' or 'text', sum_credit_balance or '', None, self.style_decimal_bold),
            ('col9', 1, 0, 'text', '', None ),
            ('col10', 1, 0, 'text', '', None ),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)

        # Header số dư cuối kỳ
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', '', None),
            ('col5', 1, 0, 'text', 'Số dư cuối kỳ', None),
            ('col6', 1, 0, 'text', '', None ),
            ('col7', 1, 0, 'text', '', None ),
            ('col8', 1, 0, 'text', '', None ),
            ('col9', 1, 0, 'number', remain_balance, None, self.style_decimal_bold ),
            ('col10', 1, 0, 'text', '', None ),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style_bold_borderall)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=self.normal_style)

        ###############
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        cell_format = _xs['wrap'] + _xs['center'] + _xs['middle']
        cell_footer_style = xlwt.easyxf(cell_format)
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(7)]
        c_specs = empty + [
            ('note1', 3, 0, 'text','Ngày .... tháng .... năm ....', None)
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
            ('col3', 1, 0, 'text', '', None),
            ('col6', 2, 16, 'text', 'Kế toán trưởng', None),
            ('col7', 1, 0, 'text', '', None),
            ('col9', 3, 0, 'text', 'Giám đốc', None),
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
            ('col5', 1, 0, 'text', '', None),
            ('col6', 2, 16, 'text', '(Ký, họ tên)', None),
            ('col8', 1, 0, 'text', '', None),
            ('col9', 3, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style=cell_footer_style)

cash_book_report_xls('report.cash_book_report_xls','account.move.line', parser=Parser)
