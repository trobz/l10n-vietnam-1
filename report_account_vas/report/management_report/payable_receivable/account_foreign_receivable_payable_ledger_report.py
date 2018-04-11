# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from datetime import datetime
from openerp.addons.report_xlsx.utils import _render  # @UnresolvedImport
from openerp.addons.report_base_vn.report import report_base_vn
from openerp.addons.report_xlsx.report_xlsx import report_xlsx
from datetime import timedelta
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
import logging
from itertools import groupby
_logger = logging.getLogger(__name__)


class account_foreign_receivable_payable_ledger_xls_parser(
        report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(account_foreign_receivable_payable_ledger_xls_parser,
              self).__init__(cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_wizard_data': self.get_wizard_data
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
        return {'name': name, 'address': address, 'vat': vat or ''}

    def get_wizard_data(self):
        result = {}
        datas = self.localcontext['data']
        if datas:
            result['fiscalyear'] = datas['form'] and datas[
                'form']['fiscalyear_id'] or False
            result['target_move'] = datas['form'] and datas[
                'form']['target_move'] or False
            result['foreign_currency_id'] = datas['form'] and datas['form'][
                'foreign_currency_id'] and datas[
                'form']['foreign_currency_id'][0] or False
            result['filter'] = 'filter' in datas[
                'form'] and datas['form']['filter'] or False
            result['partner_id'] = datas['form'] and datas['form'][
                'partner_id'] and datas['form']['partner_id'][0] or False
            result['account_type'] = datas[
                'form'] and datas['form']['result_selection']
            if datas['form']['filter'] == 'filter_date':
                result['date_from'] = datas['form']['date_from']
                result['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                result['period_from'] = datas['form']['period_from']
                result['period_to'] = datas['form']['period_to']
        return result

    def get_lines_data(self):
        """
        Get all account data
        debit_account is 11*
            credit_account
        """
        date_info = self.get_date()
        wizard_datas = self.get_wizard_data()
        previous_date_from_date = datetime.strptime(
            date_info['date_from_date'],
            DEFAULT_SERVER_DATE_FORMAT) - timedelta(days=1)
        params = {'target_move': '',
                  'previous_date_from': previous_date_from_date,
                  'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date'],
                  'currency_id': wizard_datas['foreign_currency_id'],
                  'partner': '',
                  'account_type': ''
                  }

        if wizard_datas['target_move'] == 'posted':
            params['target_move'] = "AND am.state = 'posted'"

        partner_id = wizard_datas.get('partner_id', False)
        if partner_id:
            params['partner'] = "AND aml.partner_id = %s" % partner_id

        account_type = wizard_datas.get('account_type', False)
        if account_type == 'customer':
            params['account_type'] = "AND acc.type = 'receivable'"

        elif account_type == 'supplier':
            params['account_type'] = "AND acc.type = 'payable'"

        SQL = """
                SELECT * FROM
                (
                SELECT COALESCE(aml.partner_id, -1) as partner_id,
                    NULL::date as date_created,
                    '' as serial,
                    '%(previous_date_from)s'::date as effective_date,
                    'SDDK' as description,
                    '' as counterpart_account,
                    (CASE WHEN (sum(aml.debit)-sum(aml.credit)) > 0
                    THEN (sum(aml.debit)-sum(aml.credit))
                    ELSE 0
                    END) AS debit_begin,
                    (CASE WHEN (sum(aml.debit)-sum(aml.credit)) < 0
                    THEN sum(aml.debit)-sum(aml.credit)
                    ELSE 0
                    END) AS credit_begin,
                    sum(aml.amount_currency) as amount_currency,
                    0 as debit,
                    0 as credit,
                    0 as foreign_credit,
                    0 as foreign_debit
                FROM account_move_line aml
                INNER JOIN account_move am on am.id=aml.move_id
                INNER JOIN account_account acc on acc.id=aml.account_id
                WHERE aml.date < '%(date_start)s'
                    AND aml.currency_id = '%(currency_id)s'
                    %(target_move)s
                    %(partner)s
                    %(account_type)s

                GROUP BY 1,2,3,4,5,6

                UNION ALL

                SELECT
                    COALESCE(aml.partner_id, -1),
                    aml.date_created,
                    am.name as serial,
                    aml.date as effective_date,
                    aml.name as description,
                    acc_du.code as counterpart_account,
                    0 as debit_begin,
                    0 as credit_begin,
                    0 as amount_currency,
                    aml_counterpart.credit as debit,
                    aml_counterpart.debit as credit,
                (CASE WHEN aml_counterpart.amount_currency > 0
                THEN aml_counterpart.amount_currency
                ELSE 0
                END) AS foreign_credit,
                (CASE WHEN aml_counterpart.amount_currency < 0
                THEN abs(aml_counterpart.amount_currency)
                ELSE 0
                END) AS foreign_debit
                FROM account_move_line  aml
                INNER JOIN account_move am on am.id=aml.move_id
                INNER JOIN account_account acc on acc.id=aml.account_id
                INNER JOIN account_move_line aml_counterpart
                    on aml_counterpart.counter_move_id=aml.id
                INNER JOIN account_account acc_du
                    on acc_du.id=aml_counterpart.account_id
                WHERE aml.date between '%(date_start)s' and '%(date_end)s'
                    AND aml.currency_id = '%(currency_id)s'
                    %(target_move)s
                    %(partner)s
                    %(account_type)s

                UNION ALL

                SELECT
                    COALESCE(aml.partner_id, -1),
                    aml.date_created,
                    am.name as serial,
                    aml.date as effective_date,
                    aml.name as description,
                    acc_du.code as counterpart_account,
                    0 as debit_begin,
                    0 as credit_begin,
                    0 as amount_currency,
                    aml_counterpart.credit as debit,
                    aml_counterpart.debit as credit,
                (CASE WHEN aml_counterpart.amount_currency > 0
                THEN aml_counterpart.amount_currency
                ELSE 0
                END) AS foreign_credit,
                (CASE WHEN aml_counterpart.amount_currency < 0
                THEN abs(aml_counterpart.amount_currency)
                ELSE 0
                END) AS foreign_debit
                FROM account_move_line  aml
                INNER JOIN account_move am on am.id=aml.move_id
                INNER JOIN account_account acc on acc.id=aml.account_id
                INNER JOIN account_move_line aml_counterpart
                    on aml_counterpart.id=aml.counter_move_id
                INNER JOIN account_account acc_du
                    on acc_du.id=aml_counterpart.account_id
                WHERE aml.date between '%(date_start)s' and '%(date_end)s'
                    AND aml.currency_id = '%(currency_id)s'
                    %(target_move)s
                    %(partner)s
                    %(account_type)s

                ) AS A
                ORDER BY
                    partner_id desc,
                    effective_date -- very important to order by DESC

        """ % params

        self.cr.execute(SQL)
        data = self.cr.dictfetchall()
        return data


class account_foreign_receivable_payable_ledger_xls(report_xlsx):

    def __init__(
            self, name, table, rml=False,
            parser=False, header=True, store=False):
        super(account_foreign_receivable_payable_ledger_xls, self).__init__(
            name, table, rml, parser, header, store)

    def generate_xls_report(self, _p, _xs, data, objects, workbook):

        # format
        self.style_date = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': 'dd/mm/yy',
                'text_wrap': True
            }
        )
        self.style_date_right = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': 'dd/mm/yy',
                'align': 'right',
                'text_wrap': True
            }
        )
        self.style_date_right_border = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': 'dd/mm/yy',
                'align': 'right',
                'text_wrap': True,
                'border': 1
            }
        )
        self.style_date_center_border = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': 'dd/mm/yy',
                'align': 'center',
                'text_wrap': True,
                'border': 1
            }
        )

        self.style_decimal = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': '#,##0',
                'text_wrap': True
            }
        )
        self.style_decimal_border = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': '#,##0',
                'text_wrap': True,
                'border': 1
            }
        )
        self.style_decimal_bold = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': '#,##0',
                'bold': True,
                'text_wrap': True
            }
        )
        self.style_decimal_bold_border = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': '#,##0',
                'bold': True,
                'text_wrap': True,
                'border': 1
            }
        )
        self.style_default_decimal_border = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'num_format': '#,##0.00',
                'text_wrap': True,
                'border': 1
            }
        )

        self.normal_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vjustify',
                'text_wrap': True
            }
        )
        self.normal_style_right = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'right',
                'valign': 'vjustify',
                'text_wrap': True
            }
        )
        self.normal_style_left = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left',
                'valign': 'vjustify',
                'text_wrap': True
            }
        )
        self.normal_style_center = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            }
        )
        self.normal_style_left_borderall = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left',
                'valign': 'vjustify',
                'border': 1,
                'text_wrap': True
            }
        )
        self.normal_style_right_borderall = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'right',
                'valign': 'vjustify',
                'border': 1,
                'text_wrap': True
            }
        )
        self.normal_style_italic = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vjustify',
                'italic': True,
                'text_wrap': True
            }
        )

        self.normal_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left',
                'text_wrap': True
            }
        )
        self.normal_style_borderall = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            }
        )
        self.normal_style_bold = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vjustify',
                'bold': True,
                'text_wrap': True
            }
        )
        self.normal_style_bold_left_borderall = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left',
                'valign': 'vjustify',
                'bold': True,
                'text_wrap': True,
                'border': 1
            }
        )
        self.normal_style_bold_borderall = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'border': 1,
                'text_wrap': True
            }
        )

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E',
                            'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
        self.col_specs_template = {
            'A': {
                'lines': [1, 10, _render("line.get('date_created','') and 'date' or 'text'"), _render("datetime.strptime(line.get('date_created','')[:10],'%Y-%m-%d')"), None, self.style_date_center_border],
                'totals': [1, 0, 'text', None]},

            'B': {
                'lines': [1, 12, 'text', _render("line.get('serial','') or ''"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'C': {
                'lines': [1, 10, _render("line.get('effective_date','') and 'date' or 'text'"), _render("datetime.strptime(line.get('effective_date',None)[:10],'%Y-%m-%d')"), None, self.style_date_center_border],
                'totals': [1, 0, 'text', None]},

            'D': {
                'lines': [1, 30, 'text', _render("line.get('description','') or ''"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'number', None]},

            'E': {
                'lines': [1, 6, 'text', _render("line.get('counterpart_account',None)"), None, self.normal_style_left_borderall],
                'totals': [1, 0, 'text', None]},

            'F': {
                'lines': [1, 6, 'number', _render("abs(line.get('rate',None))"), None, self.style_default_decimal_border],
                'totals': [1, 0, 'number', None]},

            'G': {
                'lines': [1, 15, 'number',  _render("abs(line.get('foreign_debit',0))"), None, self.style_default_decimal_border],
                'totals': [1, 0, 'number', None]},

            'H': {
                'lines': [1, 15, 'number', _render("abs(line.get('debit',0))"), None, self.style_decimal_border],
                'totals': [1, 0, 'number', None]},

            'I': {
                'lines': [1, 15, 'number', _render("abs(line.get('foreign_credit',0))"), None, self.style_default_decimal_border],
                'totals': [1, 0, 'number', None]},

            'J': {
                'lines': [1, 15, 'number', _render("abs(line.get('credit',0))"), None, self.style_decimal_border],
                'totals': [1, 0, 'number', None]},

            'K': {
                'lines': [1, 15, 'number', _render("abs(line.get('overbalance_foreign_debit',0))"), None, self.style_default_decimal_border],
                'totals': [1, 0, 'text', None]},

            'L': {
                'lines': [1, 15, 'number', _render("abs(line.get('overbalance_debit',0))"), None, self.style_decimal_border],
                'totals': [1, 0, 'text', None]},

            'M': {
                'lines': [1, 15, 'number', _render("abs(line.get('overbalance_foreign_credit',0))"), None, self.style_default_decimal_border],
                'totals': [1, 0, 'number', None]},

            'N': {
                'lines': [1, 15, 'number', _render("abs(line.get('overbalance_credit',0))"), None, self.style_decimal_border],
                'totals': [1, 0, 'number', None]}

        }

        NAME = ''
        report_type = data['form'] and data['form']['result_selection'] or ''
        if report_type == 'customer':
            NAME = u'MUA'
        elif report_type == 'supplier':
            NAME = u'BÁN'

        report_name = \
            u'SỔ CHI TIẾT THANH TOÁN VỚI NGƯỜI %s BẰNG NGOẠI TỆ' % NAME

        # call parent init utils.
        # set print sheet
        ws = workbook.add_worksheet(report_name[:31])
        ws.set_landscape()
        ws.set_zoom(100)
        ws.fit_to_pages(1, 1)

        row_pos = 0

        cell_address_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left',
                'bold': True,
                'text_wrap': True
            }
        )
        # Title address 1
        c_specs = [
            ('company_name', 6, 0, 'text', u'Đơn vị: %s' %
             _p.get_company()['name'], '', cell_address_style),
            ('layout', 8, 0, 'text', u'Mẫu số S14 – DNN',
             '', self.normal_style_center),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title 2
        c_specs = [
            ('company_name', 6, 0, 'text', u'Địa chỉ: %s' %
             _p.get_company()['address'], '', cell_address_style),
            ('layout', 8, 0, 'text',
                (u'(Ban hành theo QĐ số 48/2006/QĐ-BTC, '
                 u'Ngày 14/09/2006 của Bộ trưởng BTC)'),
             '', self.normal_style_center),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('company_name', 6, 0, 'text', u'MST: %s' %
             _p.get_company()['vat'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [
            ('col1', 1, 10, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
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
        cell_title_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'text_wrap': True,
                'font_size': 18
            })
        ws.set_row(row_pos, 22)
        c_specs = [
            ('payment_journal', 14, 0, 'text', report_name, None,
                cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Loại ngoại tệ: USD
        currency_name = data['form'] and data['form'][
            'foreign_currency_id'] and data['form'][
            'foreign_currency_id'][1] or ''
        c_specs = [
            ('from_to', 14, 0, 'text', u'Loại ngoại tệ: %s' % currency_name)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Title "Từ .... Đến ...."
        c_specs = [
            ('from_to', 14, 0, 'text', u'Từ %s đến %s' % (_p.get_date().get(
                'date_from', '.......'), _p.get_date().get(
                'date_to', '.......')))
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

        h1_c_specs = [
            ('col1', 1, 10, 'text', u'Ngày tháng ghi sổ', None),
            ('col2', 2, 22, 'text', u'Chứng Từ', None),
            ('col3', 1, 40, 'text', u'Diễn giải', None),
            ('col4', 1, 15, 'text', u'TK đối ứng', None),
            ('col5', 1, 15, 'text', u'Tỉ giá hối đoái', None),
            ('col6', 4, 60, 'text', u'Số phát sinh', None),
            ('col7', 4, 60, 'text', u'Số dư', None),
        ]
        row_data = self.xls_row_template(
            h1_c_specs, [x[0] for x in h1_c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_bold_borderall, set_column_size=True)

        # Header Title 2
        h2_c_specs = [
            ('col1', 1, 10, 'text', '', None),
            ('col2', 1, 12, 'text', u'Số hiệu', None),
            ('col3', 1, 10, 'text', u'Ngày tháng', None),
            ('col4', 1, 40, 'text', '', None),
            ('col5', 1, 15, 'text', '', None),
            ('col6', 1, 15, 'text', '', None),
            ('col7', 2, 30, 'text', u'Nợ', None),
            ('col8', 2, 30, 'text', u'Có', None),
            ('col9', 2, 30, 'text', u'Nợ', None),
            ('col10', 2, 30, 'text', u'Có', None),
        ]
        row_data = self.xls_row_template(
            h2_c_specs, [x[0] for x in h2_c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.merge_range(row_title_body_pos, 0, row_title_body_pos + 1,
                       0, u'Ngày tháng ghi sổ',
                       self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 3, row_title_body_pos + 1,
                       3, u'Diễn giải', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 4, row_title_body_pos + 1,
                       4, u'TK đối ứng', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 5, row_title_body_pos + 1,
                       5, u'Tỉ giá hối đoái', self.normal_style_bold_borderall)

        # Header Title 3
        h3_c_specs = [
            ('col1', 1, 10, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 10, 'text', '', None),
            ('col4', 1, 40, 'text', '', None),
            ('col5', 1, 15, 'text', '', None),
            ('col6', 1, 15, 'text', '', None),
            ('col7', 1, 15, 'text', u'Ngoại tệ', None),
            ('col8', 1, 15, 'text', u'Quy ra VNĐ', None),
            ('col9', 1, 15, 'text', u'Ngoại tệ', None),
            ('col10', 1, 15, 'text', u'Quy ra VNĐ', None),
            ('col11', 1, 15, 'text', u'Ngoại tệ', None),
            ('col12', 1, 15, 'text', u'Quy ra VNĐ', None),
            ('col13', 1, 15, 'text', u'Ngoại tệ', None),
            ('col14', 1, 15, 'text', u'Quy ra VNĐ', None),
        ]
        row_data = self.xls_row_template(
            h3_c_specs, [x[0] for x in h3_c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_bold_borderall, set_column_size=True)

        # merge cell
        ws.merge_range(row_title_body_pos, 0, row_title_body_pos + 2,
                       0, u'Ngày tháng ghi sổ',
                       self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 1, row_title_body_pos + 1,
                       1, u'Số hiệu', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 2, row_title_body_pos + 1,
                       2, u'Ngày tháng', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 3, row_title_body_pos + 2,
                       3, u'Diễn giải', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 4, row_title_body_pos + 2,
                       4, u'TK đối ứng', self.normal_style_bold_borderall)
        ws.merge_range(row_title_body_pos, 5, row_title_body_pos + 2,
                       5, u'Tỉ giá hối đoái', self.normal_style_bold_borderall)

        # account move lines
        get_lines_data = _p.get_lines_data()

        # @UnusedVariable
        for partner_id, group_datas in groupby(
                get_lines_data, lambda line: line['partner_id']):
            #
            group_datas = list(group_datas)
            overbalance_amount = 0
            foreign_overbalance_amount = 0
            sum_foreign_debit = 0
            sum_foreign_credit = 0
            sum_debit = 0
            sum_credit = 0
            end_debit = 0
            end_credit = 0
            end_foreign_debit = 0
            end_foreign_credit = 0
            debit_begin = 0
            credit_begin = 0

            row_pos = self.write_partner(
                ws, row_pos, partner_id, row_style=self.normal_style_borderall)

            if group_datas[0].get('description') == 'SDDK':
                line = group_datas[0]
                debit_begin = line['debit_begin']
                credit_begin = line['credit_begin']
                overbalance_amount = debit_begin - abs(credit_begin)
                foreign_overbalance_amount = line['amount_currency']
                if foreign_overbalance_amount > 0:
                    begining_c_specs = self.gen_sddk_c_cpecs(
                        debit_begin, credit_begin,
                        foreign_overbalance_amount, 0)
                else:
                    begining_c_specs = self.gen_sddk_c_cpecs(
                        debit_begin, credit_begin, 0,
                        foreign_overbalance_amount)
                row_pos = self.write_sddk(
                    ws, row_pos, begining_c_specs,
                    row_style=self.normal_style_borderall)
                group_datas = group_datas[1:]

                #####

                # Sum total and ending
                end_foreign_debit += foreign_overbalance_amount > 0 and \
                    foreign_overbalance_amount or 0
                end_foreign_credit += foreign_overbalance_amount < 0 and \
                    foreign_overbalance_amount or 0

                end_debit += debit_begin
                end_credit += credit_begin

            else:
                begining_c_specs = self.gen_sddk_c_cpecs(0, 0, 0, 0)
                row_pos = self.write_sddk(
                    ws, row_pos, begining_c_specs,
                    row_style=self.normal_style_borderall)

                # Sum total and ending
                end_foreign_debit += foreign_overbalance_amount > 0 and \
                    foreign_overbalance_amount or 0
                end_foreign_credit += foreign_overbalance_amount < 0 and \
                    foreign_overbalance_amount or 0

                end_debit += debit_begin
                end_credit += credit_begin

            for line in group_datas:
                row_pos, overbalance_amount, foreign_overbalance_amount, \
                    end_debit, end_credit, sum_foreign_debit,\
                    sum_foreign_credit, sum_debit, sum_credit, \
                    end_foreign_debit, end_foreign_credit = \
                    self.write_lines(
                        ws, row_pos, line, overbalance_amount,
                        foreign_overbalance_amount, end_debit, end_credit,
                        sum_foreign_debit, sum_foreign_credit, sum_debit,
                        sum_credit, end_foreign_debit, end_foreign_credit,
                        row_style=self.normal_style_borderall
                    )

            cps_c_specs, ending_c_specs = self.gen_cps_ending_c_cpecs(
                sum_foreign_debit, sum_debit,
                sum_foreign_credit, sum_credit,
                end_foreign_debit, end_debit,
                end_foreign_credit, end_credit
            )
            # Cộng phát sinh
            row_data = self.xls_row_template(
                cps_c_specs, [x[0] for x in cps_c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data,
                row_style=self.normal_style_borderall, set_column_size=True)

            # Số dư cuối kỳ
            row_data = self.xls_row_template(
                ending_c_specs, [x[0] for x in ending_c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data,
                row_style=self.normal_style_borderall, set_column_size=True)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 0, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        ###############
        cell_footer_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            }
        )

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = empty + [
            ('node1', 3, 0, 'text', u'Ngày .... tháng .... năm ....', None)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        cell_footer_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'bold': True
            }
        )

        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', u'Người ghi sổ', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 1, 0, 'text', '', None),
            ('col7', 3, 0, 'text', u'Kế toán trưởng', None),
            ('col10', 1, 0, 'text', '', None),
            ('col11', 1, 0, 'text', '', None),
            ('col12', 3, 0, 'text', u'Giám đốc', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        ###############
        cell_footer_style = workbook.add_format(
            {
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'center',
                'italic': True
            }
        )
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(11)]
        c_specs = [
            ('col1', 1, 0, 'text', '', None),
            ('col2', 1, 0, 'text', '', None),
            ('col3', 1, 0, 'text', '', None),
            ('col4', 1, 0, 'text', u'(Ký, họ tên)', None),
            ('col5', 1, 0, 'text', '', None),
            ('col6', 1, 0, 'text', '', None),
            ('col7', 3, 0, 'text', u'(Ký, họ tên)', None),
            ('col10', 1, 0, 'text', '', None),
            ('col11', 1, 0, 'text', '', None),
            ('col12', 3, 0, 'text', u'(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)

        workbook.close()

    def write_partner(self, ws, row_pos, partner_id, row_style):
        if partner_id:
            partner_obj = self.pool['res.partner']
            partner_data = partner_obj.read(
                self.cr, self.uid, partner_id, ['name'])
            partner_name = partner_data and partner_data['name'] or ''
            partner_c_specs = [
                ('col1', 4, 12, 'text', partner_name, None,
                 self.normal_style_bold_left_borderall),
                ('col5', 1, 12, 'text', '', None),
                ('col6', 1, 12, 'text', '', None),
                ('col7', 1, 12, 'text', '', None),
                ('col8', 1, 12, 'text', '', None),
                ('col9', 1, 12, 'text', '', None),
                ('col10', 1, 12, 'text', '', None),
                ('col11', 1, 12, 'text', '', None),
                ('col12', 1, 12, 'text', '', None),
                ('col13', 1, 12, 'text', '', None),
                ('col14', 1, 12, 'text', '', None),
            ]
            row_data = self.xls_row_template(
                partner_c_specs, [x[0] for x in partner_c_specs])
            row_pos = self.xls_write_row(
                ws, row_pos, row_data, row_style, set_column_size=True)
        return row_pos

    def write_sddk(self, ws, row_pos, begining_c_specs, row_style):
        # Số dư đầu kỳ
        row_data = self.xls_row_template(
            begining_c_specs, [x[0] for x in begining_c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style, set_column_size=True)
        return row_pos

    def calculate_overbalance(self, overbalance_amount, debit, credit):
        overbalance_amount = overbalance_amount + debit - abs(credit)
        overbalance_debit = overbalance_credit = end_debit = end_credit = 0
        if overbalance_amount > 0:
            overbalance_debit = end_debit = overbalance_amount
            overbalance_credit = end_credit = 0
        elif overbalance_amount < 0:
            overbalance_credit = end_credit = overbalance_amount
            overbalance_debit = end_debit = 0
        return overbalance_amount, overbalance_debit, overbalance_credit, end_debit, end_credit

    def write_lines(
            self, ws, row_pos, line, overbalance_amount,
            foreign_overbalance_amount, end_debit, end_credit,
            sum_foreign_debit, sum_foreign_credit, sum_debit,
            sum_credit, end_foreign_debit, end_foreign_credit, row_style):
        #  Write lines
        debit = line['debit']
        credit = line['credit']
        foreign_debit = line['foreign_debit']
        foreign_credit = line['foreign_credit']

        # Calculate overbalance
        overbalance_amount, overbalance_debit, \
            overbalance_credit, end_debit, end_credit =\
            self.calculate_overbalance(overbalance_amount, debit, credit)

        foreign_overbalance_amount, overbalance_foreign_debit, \
            overbalance_foreign_credit, end_foreign_debit, \
            end_foreign_credit = self.calculate_overbalance(
                foreign_overbalance_amount,
                foreign_debit,
                foreign_credit
            )

        line['overbalance_debit'] = overbalance_debit
        line['overbalance_credit'] = overbalance_credit

        line['overbalance_foreign_debit'] = overbalance_foreign_debit
        line['overbalance_foreign_credit'] = overbalance_foreign_credit

        # SUm foreign
        sum_foreign_debit += foreign_debit
        sum_foreign_credit += foreign_credit

        sum_debit += debit
        sum_credit += credit

        # Exchange Rate
        rate = 1
        if debit > 0 and foreign_debit != 0:
            rate = debit / foreign_debit
        elif credit > 0 and foreign_credit != 0:
            rate = credit / foreign_credit

        line['rate'] = rate
        # write
        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'lines'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data,
            row_style=self.normal_style_borderall, set_column_size=True)
        return row_pos, overbalance_amount, foreign_overbalance_amount,\
            end_debit, end_credit, sum_foreign_debit, sum_foreign_credit,\
            sum_debit, sum_credit, end_foreign_debit, end_foreign_credit

    def gen_sddk_c_cpecs(
            self, debit_begin, credit_begin,
            foreign_debit_begin, foreign_credit_begin):
        begining_c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', u'Số dư đầu kỳ',
             None, self.normal_style_bold_left_borderall),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 12, 'text', '', None),
            ('col8', 1, 12, 'text', '', None),
            ('col9', 1, 12, 'text', '', None),
            ('col10', 1, 12, 'text', '', None),
            ('col11', 1, 12, 'number', abs(foreign_debit_begin),
             None, self.style_default_decimal_border),
            ('col12', 1, 12, 'number', abs(debit_begin),
             None, self.style_decimal_border),
            ('col13', 1, 12, 'number', abs(foreign_credit_begin),
             None, self.style_default_decimal_border),
            ('col14', 1, 12, 'number', abs(credit_begin),
             None, self.style_decimal_border),
        ]
        return begining_c_specs

    def gen_cps_ending_c_cpecs(
            self, sum_foreign_debit, sum_debit, sum_foreign_credit, sum_credit,
            end_foreign_debit, end_debit, end_foreign_credit, end_credit):
        cps_c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', u'Cộng phát sinh',
             None, self.normal_style_bold_left_borderall),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 12, 'number', abs(sum_foreign_debit),
             None, self.style_default_decimal_border),
            ('col8', 1, 12, 'number', abs(sum_debit),
             None, self.style_decimal_border),
            ('col9', 1, 12, 'number', abs(sum_foreign_credit),
             None, self.style_default_decimal_border),
            ('col10', 1, 12, 'number', abs(sum_credit),
             None, self.style_decimal_border),
            ('col11', 1, 12, 'text', '', None),
            ('col12', 1, 12, 'text', '', None),
            ('col13', 1, 12, 'text', '', None),
            ('col14', 1, 12, 'text', '', None),
        ]

        ending_c_specs = [
            ('col1', 1, 12, 'text', '', None),
            ('col2', 1, 12, 'text', '', None),
            ('col3', 1, 12, 'text', '', None),
            ('col4', 1, 12, 'text', u'Số dư cuối kỳ',
             None, self.normal_style_bold_left_borderall),
            ('col5', 1, 12, 'text', '', None),
            ('col6', 1, 12, 'text', '', None),
            ('col7', 1, 12, 'text', '', None),
            ('col8', 1, 12, 'text', '', None),
            ('col9', 1, 12, 'text', '', None),
            ('col10', 1, 12, 'text', '', None),
            ('col11', 1, 12, 'number', abs(end_foreign_debit),
             None, self.style_default_decimal_border),
            ('col12', 1, 12, 'number', abs(end_debit),
             None, self.style_decimal_border),
            ('col13', 1, 12, 'number', abs(end_foreign_credit),
             None, self.style_default_decimal_border),
            ('col14', 1, 12, 'number', abs(end_credit),
             None, self.style_decimal_border),
        ]
        return cps_c_specs, ending_c_specs


account_foreign_receivable_payable_ledger_xls(
    'report.foreign_receivable_payable_ledger_report',
    'account.move.line',
    parser=account_foreign_receivable_payable_ledger_xls_parser)
