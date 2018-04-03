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


import xlwt
from datetime import datetime
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from openerp.addons.report_base_vn.report import report_base_vn
from openerp.addons.report_account_vas.report import report_xls_utils
import logging
_logger = logging.getLogger(__name__)


class GeneralDetailPayableReceivableBalanceParser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(GeneralDetailPayableReceivableBalanceParser, self).__init__(
            cr, uid, name, context=context
        )
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
            'get_lines_data': self.get_lines_data,
            'get_wizard_data': self.get_wizard_data,
            'is_purchases_journal': True,
            'get_company': self.get_company,
            'get_partner_info': self.get_partner_info
        })

    def get_company(self):
        res = self.pool['res.users'].browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
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
        return {'name': name, 'address': address, 'vat': vat}

    def get_partner_info(self, partner_id, context=None):
        partner_rec = self.pool['res.partner'].browse(
            self.cr, self.uid, partner_id,
            context=context
        )
        res = {
            'ref': partner_rec.ref or '',
            'name': partner_rec.name or ''
        }

        return res

    def get_data(self):
        return self.localcontext['data']['form']

    def get_wizard_data(self):
        res = {}
        data = self.get_data()

        res.update({
            'from': data['date_from'],
            'to': data['date_to']
        })
        if data['target_move'] == 'posted':
            res['target_move'] = ('posted', '1')
        elif data['target_move'] == 'all':
            res['target_move'] = ('posted', 'draft')

        res['partner_id'] = data['partner_id'][0]
        res['account_id'] = data['account_id'][0]
        res['account_name'] = data['account_id'][1]
        res['account_type'] = data['account_type']
        res['journal_ids'] = tuple(data['journal_ids'] or [-1, -1])

        return res

    def get_lines_data(self, start_date, to_date, type, state, partner_id,
                       journal_ids, account_id):
        """
        """
        params = {
            'date_from': start_date,
            'date_to': to_date,
            'type': type,
            'state': state,
            'partner_id': partner_id,
            'account_id': account_id,
            'journal_ids': journal_ids
        }

        SQL = """

            SELECT
                '' as serial,
                '%(date_from)s'::date as effective_date,
                'SDDK' as description,
                '' as counterpart_account,
                (
                    CASE WHEN (sum(aml.debit)-sum(aml.credit)) > 0
                    THEN (sum(aml.debit)-sum(aml.credit))
                    ELSE 0
                     END
                ) AS debit,
                (
                    CASE WHEN (sum(aml.debit)-sum(aml.credit)) < 0
                    THEN abs(sum(aml.debit)-sum(aml.credit))
                    ELSE 0
                     END
                ) AS credit
            FROM account_move_line aml
            INNER JOIN account_move am on am.id=aml.move_id
            INNER JOIN account_account acc on acc.id=aml.account_id
            WHERE aml.date < '%(date_from)s'
                and am.state in %(state)s
                and acc.internal_type = '%(type)s'
                and acc.id = %(account_id)s
                and aml.partner_id=%(partner_id)s
                and aml.journal_id IN %(journal_ids)s
            GROUP BY 1,2,3,4

            UNION ALL

            SELECT * FROM (
                SELECT
                    am.name as serial,
                    aml.date as effective_date,
                    aml.name as description,
                    acc_du.code as counterpart_account,
                    aml_counterpart.credit as debit,
                    aml_counterpart.debit as credit
                FROM account_move_line  aml
                INNER JOIN account_move am on am.id=aml.move_id
                INNER JOIN account_account acc on acc.id=aml.account_id
                INNER JOIN account_move_line aml_counterpart on aml_counterpart.counter_move_id=aml.id
                INNER JOIN account_account acc_du on acc_du.id=aml_counterpart.account_id
                WHERE aml.date between '%(date_from)s' and '%(date_to)s'
                    and am.state in %(state)s
                    and acc.internal_type='%(type)s'
                    and acc.id = %(account_id)s
                    and aml.partner_id=%(partner_id)s
                    and aml.journal_id IN %(journal_ids)s

                UNION ALL

                SELECT
                    am.name as serial,
                    aml.date as effective_date,
                    aml.name as description,
                    acc_du.code as counterpart_account,
                    aml_counterpart.credit as debit,
                    aml_counterpart.debit as credit
                FROM account_move_line  aml
                INNER JOIN account_move am on am.id=aml.move_id
                INNER JOIN account_account acc on acc.id=aml.account_id
                INNER JOIN account_move_line aml_counterpart on aml_counterpart.id=aml.counter_move_id
                INNER JOIN account_account acc_du on acc_du.id=aml_counterpart.account_id
                WHERE aml.date between '%(date_from)s' and '%(date_to)s'
                    and am.state in %(state)s
                    and acc.internal_type='%(type)s'
                    and acc.id = %(account_id)s
                    and aml.partner_id=%(partner_id)s
                    and aml.journal_id IN %(journal_ids)s

            ) in_period

            ORDER BY effective_date,serial,description

        """ % params

        self.cr.execute(SQL)
        data = self.cr.dictfetchall()
        return data


class GeneralDetailPayableReceivableBalanceXLS(report_xls_utils.generic_report_xls_base):

    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(GeneralDetailPayableReceivableBalanceXLS, self).__init__(
            name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

        # XLS Template
        self.wanted_list = ['A', 'B', 'C', 'D', 'E', 'F']
        self.col_specs_template = {
            'A': {
                'lines': [1, 12, _render("line.get('effective_date','') and 'date' or 'text'"),
                          _render(
                              "datetime.strptime(line.get('effective_date','')[:10],'%Y-%m-%d')"),
                          None, self.style_date_right],
                'totals': [1, 12, 'text', None]},

            'B': {
                'lines': [1, 20, 'text', _render("line.get('serial','')"),
                          None, self.normal_style_left_borderall],
                'totals': [1, 12, 'text', None]},

            'C': {
                'lines': [1, 30, 'text', _render("line.get('description',0)"),
                          None, self.normal_style_left_borderall],
                'totals': [1, 20, 'text', u'Tổng Cộng', None,
                           self.normal_style_bold_borderall]},

            'D': {
                'lines': [1, 15, 'text', _render("line.get('counterpart_account',0)"),
                          None, self.normal_style_right_borderall],
                'totals': [1, 15, 'text', None]},

            'E': {
                'lines': [1, 15, _render("line.get('debit') and 'number' or 'text'"), _render("line.get('debit',0)"),
                          None, self.style_decimal],
                'totals': [1, 15, _render("sum_debit and 'number' or 'text'"), _render('sum_debit') or '', None,
                           self.style_decimal_bold]},

            'F': {
                'lines': [1, 15, _render("line.get('credit') and 'number' or 'text'"), _render("line.get('credit',0)"),
                          None, self.style_decimal],
                'totals': [1, 15, _render("sum_credit and 'number' or 'text'"), _render('sum_credit') or '', None,
                           self.style_decimal_bold]},
        }

    def gen_c_cpecs(self, name, debit, credit):
        begining_c_specs = [
            ('col1', 1, 12, 'text', None),
            ('col2', 1, 20, 'text', None),
            ('col3', 1, 30, 'text', name, None,
             self.normal_style_bold_borderall),
            ('col4', 1, 15, 'text', None),
            ('col5', 1, 15, debit and 'number' or 'text', debit, None,
             self.style_decimal_bold),
            ('col6', 1, 15, credit and 'number' or 'text', credit, None,
             self.style_decimal_bold),
        ]
        return begining_c_specs

    def write_one_line(self, ws, row_pos, c_specs, row_style):
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data, row_style,
                                     set_column_size=True)
        return row_pos

    def generate_xls_header(self, _p, _xs, data, objects, wb, ws, row_pos,
                            report_name, date_start, date_to, partner_id):
        """

        """
        partner_info = _p.get_partner_info(partner_id)

        cell_address_style = self.get_cell_style(['bold', 'wrap', 'left'])
        # Title address 1
        c_specs = [
            ('company_name', 6, 0, 'text', u'Đơn vị báo cáo: %s'
             % _p.get_company()['name'], '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style)

        # Title 2
        c_specs = [
            ('dc', 6, 0, 'text', u'Địa chỉ: %s' % _p.get_company()['address'],
             '', cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style)

        # Title 3
        c_specs = [
            ('mst', 6, 0, 'text', u'MST: %s' % _p.get_company()['vat'], '',
             cell_address_style),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 20, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style)

        # Title "SỔ NHẬT KÝ MUA HÀNG"
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 256 * 2
        cell_title_style = self.get_cell_style(['bold', 'wrap', 'center',
                                                'middle', 'fontsize_350'])
        c_specs = [
            ('payment_journal', 6, 0, 'text', report_name, None,
             cell_title_style)
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style)

        # Ma khach hang
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('customer_ref', 6, 0, 'text', u'Mã khách hàng: %s'
                % partner_info['ref'] or '')
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style_italic)

        # Customer name
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('customer_ref', 6, 0, 'text', u'Tên khách hàng: %s'
                % partner_info['name'] or '')
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style_italic)

        # Title "Từ .... Đến ...."
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('from_to', 6, 0, 'text', u'Từ %s đến %s' %
             (datetime.strptime(date_start, '%Y-%m-%d').strftime('%d-%m-%Y') or '.........',
              datetime.strptime(date_to, '%Y-%m-%d').strftime('%d-%m-%Y') or '..........'))
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # add account_name
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [(
            'account_id', 6, 0, 'text',
            u'Tài khoản: %s' % _p.get_wizard_data()['account_name']
        )]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style_italic)

        # Add 1 empty line
        c_specs = [('empty%s' % x, 1, 20, 'text', '') for x in range(6)]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.normal_style)

        # Header Title 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 2, 17, 'text', 'Chứng Từ', None),
            ('col2', 1, 34, 'text', 'Diễn Giải', None),
            ('col3', 1, 20, 'text', 'Tài khoản đối ứng', None),
            ('col4', 1, 40, 'text', 'Nợ', None),
            ('col5', 1, 40, 'text', 'Có', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style_bold_borderall,
                                     set_column_size=True)

        # Header Title 1
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 300
        c_specs = [
            ('col1', 1, 12, 'text', 'Ngày Tháng', None),
            ('col2', 1, 20, 'text', 'Số hiệu', None),
            ('col3', 1, 30, 'text', '', None),
            ('col4', 1, 15, 'text', '', None),
            ('col5', 1, 15, 'text', '', None),
            ('col6', 1, 15, 'text', '', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(ws, row_pos, row_data,
                                     row_style=self.normal_style_bold_borderall,
                                     set_column_size=True)

        ws.write_merge(row_pos - 2, row_pos - 1, 2, 2, u'Diễn Giải',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_pos - 2, row_pos - 1, 3, 3, u'Tài Khoản Đối Ứng',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_pos - 2, row_pos - 1, 4, 4, u'Nợ',
                       self.normal_style_bold_borderall)
        ws.write_merge(row_pos - 2, row_pos - 1, 5, 5, u'Có',
                       self.normal_style_bold_borderall)

        return row_pos

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        # Wizard information
        wizard_info = _p.get_wizard_data()

        report_name = u'SỔ CHI TIẾT CÔNG NỢ PHẢI THU'
        if wizard_info['account_type'] == 'payable':
            report_name = u'SỔ CHI TIẾT CÔNG NỢ PHẢI TRẢ'

        # call parent init utils.
        # set print sheet
        ws = super(GeneralDetailPayableReceivableBalanceXLS, self).generate_xls_report(
            _p, _xs, data, objects, wb, report_name)
        row_pos = self.generate_xls_header(_p, _xs, data, objects, wb, ws, 0,
                                           report_name, wizard_info['from'],
                                           wizard_info['to'],
                                           wizard_info['partner_id'])
        MAX_ROW = 65500
        count = 1

        # account move lines
        lines_data = _p.get_lines_data(wizard_info['from'], wizard_info['to'],
                                       wizard_info['account_type'],
                                       wizard_info['target_move'],
                                       wizard_info['partner_id'],
                                       wizard_info['journal_ids'],
                                       wizard_info['account_id'])

        sum_debit = 0
        sum_credit = 0
        flag_write_sddk = False
        credit_sddk = 0
        debit_sddk = 0
        for line in lines_data:  # @UnusedVariable
            flag_not_write = False
            sddk = line.get('description', '')
            if not flag_write_sddk:
                if sddk == 'SDDK':
                    debit = line['debit']
                    credit = line['credit']
                    credit_sddk = credit
                    debit_sddk = debit
                    begining_c_specs = self.gen_c_cpecs(u'Số dư đầu kỳ', debit,
                                                        credit)
                    row_pos = self.write_one_line(ws, row_pos, begining_c_specs,
                                                  row_style=self.normal_style_borderall)
                    flag_write_sddk = True
                    flag_not_write = True

                else:
                    begining_c_specs = self.gen_c_cpecs(u'Số dư đầu kỳ', 0, 0)
                    row_pos = self.write_one_line(ws, row_pos, begining_c_specs,
                                                  row_style=self.normal_style_borderall)
                    flag_write_sddk = True

            if not flag_not_write:
                if row_pos > MAX_ROW:
                    ws.flush_row_data()
                    ws, row_pos = self.generate_worksheet(
                        _p, _xs, data, objects, wb, report_name, count)
                    row_pos = self.generate_xls_header(
                        _p, _xs, data, objects, wb, ws, row_pos, report_name)
                    count += 1
                c_specs = map(
                    lambda x: self.render(x, self.col_specs_template, 'lines'), self.wanted_list)
                row_data = self.xls_row_template(
                    c_specs, [x[0] for x in c_specs])
                row_pos = self.xls_write_row(
                    ws, row_pos, row_data, row_style=self.normal_style_borderall)

                sum_debit += line.get('debit', 0)
                sum_credit += line.get('credit', 0)

        # Totals
        ws.row(row_pos).height_mismatch = True
        ws.row(row_pos).height = 450

        c_specs = map(lambda x: self.render(
            x, self.col_specs_template, 'totals'), self.wanted_list)
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=self.style_decimal_bold)

        # Write SDCK
        sdck = (debit_sddk - credit_sddk) + (sum_debit - sum_credit)
        if sdck > 0:
            ending_c_specs = self.gen_c_cpecs(u'Số dư cuối kỳ', sdck, 0)
            row_pos = self.write_one_line(ws, row_pos, ending_c_specs,
                                          row_style=self.style_decimal_bold)
        else:
            ending_c_specs = self.gen_c_cpecs(u'Số dư cuối kỳ', 0, abs(sdck))
            row_pos = self.write_one_line(ws, row_pos, ending_c_specs,
                                          row_style=self.style_decimal_bold)

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
        empty = [('empty%s' % x, 1, 0, 'text', '') for x in range(4)]
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

        c_specs = [
            ('col2', 2, 0, 'text', 'Người ghi sổ', None),
            ('col6', 2, 0, 'text', 'Kế toán trưởng', None),
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
        c_specs = [
            ('col2', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col6', 2, 0, 'text', '(Ký, họ tên)', None),
            ('col9', 2, 0, 'text', '(Ký, họ tên, đóng dấu)', None),
        ]
        row_data = self.xls_row_template(c_specs, [x[0] for x in c_specs])
        row_pos = self.xls_write_row(
            ws, row_pos, row_data, row_style=cell_footer_style)


GeneralDetailPayableReceivableBalanceXLS(
    'report.general_detail_receivable_payable_balance_xls',
    'account.move',
    parser=GeneralDetailPayableReceivableBalanceParser
)

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
