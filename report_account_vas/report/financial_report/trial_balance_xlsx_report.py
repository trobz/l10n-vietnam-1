# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from collections import OrderedDict
from datetime import datetime
import logging

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from openerp.api import Environment
from openerp import SUPERUSER_ID

_logger = logging.getLogger(__name__)


class TrialBalanceReportXlsx(ReportXlsx):

    def create_xlsx_report(self, ids, data, report):
        self.env = Environment(self.env.cr, SUPERUSER_ID, self.env.context)
        return super(TrialBalanceReportXlsx, self).create_xlsx_report(
            ids, data, report)

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]

        # choose responsible user of company on wizard
#         user_id = self.object.company_id.resp_user_id.id
#         self.object = self.object.sudo(user=user_id)
#         self.env = Environment(self.env.cr, user_id, self.env.context)

        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self.sheet.hide_gridlines(2)
        self.setup_config()

        # generate header
        self.generate_company_info()

        # generate main content
        self.generate_report_title()
        self.generate_content_header()
        self.generate_main_content()

        self.generate_footer_content()

    def setup_config(self):
        self._set_default_format()

    def _define_formats(self, workbook):
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        format_config = {
            'font_name': 'Arial',
            'font_size': 10,
            'valign': 'vcenter',
            'text_wrap': True,
        }
        self.format_default = workbook.add_format(format_config)

        format_bold = format_config.copy()
        format_bold.update({
            'bold': True,
        })
        self.format_bold = workbook.add_format(format_bold)

        format_bold_center = format_bold.copy()
        format_bold_center.update({
            'align': 'center',
        })
        self.format_bold_center = workbook.add_format(format_bold_center)

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
            'num_format': '#,##0',
        })
        self.format_table = workbook.add_format(format_table)

        format_table_date = format_config.copy()
        format_table_date.update({
            'border': True,
            'align': 'right',
            'num_format': 'dd/mm/yyyy'
        })
        self.format_table_date = workbook.add_format(format_table_date)

        format_table_center = format_table.copy()
        format_table_center.update({
            'align': 'center',
            'num_format': '@',
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

    def _set_default_format(self):
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(4, 20)

        self.sheet.set_column('A:A', 15)
        self.sheet.set_column('B:B', 30)
        self.sheet.set_column('C:C', 15)
        self.sheet.set_column('D:D', 15)
        self.sheet.set_column('E:E', 15)
        self.sheet.set_column('F:F', 15)
        self.sheet.set_column('G:G', 15)
        self.sheet.set_column('H:H', 15)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()
        info = u'{0}\n{1}\n{2}'.format(
            u'Đơn vị báo cáo: ' + company_info['name'],
            u'Địa chỉ: ' + company_info['address'],
            u'MST: '
        )
        self.sheet.merge_range(
            'A1:B3', info,
            self.format_company_info)

        self.sheet.merge_range(
            'F1:H3', u'Mẫu số S06 – DN\n'
                     u'(Ban hành theo Thông tư số 200/2014/TT-BTC, '
                     u'Ngày 22/12/2014 của Bộ Tài chính)',
            self.format_bold_center)

        return info

    def generate_report_title(self):
        self.sheet.merge_range(
            'A{r}:H{r}'.format(r=4),
            u'BẢNG CÂN ĐỐI SỐ PHÁT SINH',
            self.format_report_title)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A{r}:H{r}'.format(r=5),
            u'Từ %s đến %s' % (date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.merge_range(
            'G{r}:H{r}'.format(r=7),
            u'Đơn vị tính: %s' % self.get_currency_unit(),
            self.format_uom)

    def get_currency_unit(self):
        return self.env.user.company_id.currency_id.name

    def get_date_info(self):
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {
            'date_from': date_from,
            'date_to': date_to
        }

    def convert_date_format(self, date_str):
        return datetime.strptime(date_str, DF).strftime('%d/%m/%Y')

    def generate_content_header(self):
        self.ws_row = 8
        self.sheet.merge_range(
            'A{rs}:A{re}'.format(
                rs=self.ws_row, re=self.ws_row + 1), u'Số tài khoản',
            self.format_table_bold_center)
        self.sheet.merge_range(
            'B{rs}:B{re}'.format(
                rs=self.ws_row, re=self.ws_row + 1), u'Tên tài khoản',
            self.format_table_bold_center)

        self.sheet.merge_range(
            'C{r}:D{r}'.format(r=self.ws_row),
            u'Số dư đầu kỳ',
            self.format_table_bold_center)
        self.sheet.write(
            'C{r}'.format(r=self.ws_row + 1), u'Nợ',
            self.format_table_bold_center)
        self.sheet.write(
            'D{r}'.format(r=self.ws_row + 1), u'Có',
            self.format_table_bold_center)

        self.sheet.merge_range(
            'E{r}:F{r}'.format(r=self.ws_row),
            u'Phát sinh trong kỳ',
            self.format_table_bold_center)
        self.sheet.write(
            'E{r}'.format(r=self.ws_row + 1), u'Nợ',
            self.format_table_bold_center)
        self.sheet.write(
            'F{r}'.format(r=self.ws_row + 1), u'Có',
            self.format_table_bold_center)

        self.sheet.merge_range(
            'G{r}:H{r}'.format(r=self.ws_row),
            u'Số dư cuối kỳ',
            self.format_table_bold_center)
        self.sheet.write(
            'G{r}'.format(r=self.ws_row + 1), u'Nợ',
            self.format_table_bold_center)
        self.sheet.write(
            'H{r}'.format(r=self.ws_row + 1), u'Có',
            self.format_table_bold_center)

        # Start to write data
        self.ws_row += 2

    def get_params_for_query(self):
        return {
            'date_from': self.object.date_from,
            'date_to': self.object.date_to,
            'journal_ids': 'IN {j}'.format(
                j=tuple(self.object.journal_ids.ids + [-1, -1])),
            'disp_acc': self.object.display_account,
            'target_move': ('posted', '1')
            if self.object.target_move == 'posted'
            else ('posted', 'draft'),
            'company_id': self.object.company_id.id
            if self.object.company_id else self.env.user.company_id.id
        }

    def _write_line(self, k, v, row=1, data=None):
        if v['level'] == 1:
            fm = self.format_table_bold
        else:
            fm = self.format_table

        self.sheet.write('A{r}'.format(r=row), k, fm)
        self.sheet.write('B{r}'.format(r=row), v['name'], fm)
        self.sheet.write('C{r}'.format(r=row), data['dr_begin_period'], fm)
        self.sheet.write('D{r}'.format(r=row), data['cr_begin_period'], fm)
        self.sheet.write('E{r}'.format(r=row), data['dr_in_period'], fm)
        self.sheet.write('F{r}'.format(r=row), data['cr_in_period'], fm)
        self.sheet.write('G{r}'.format(r=row), data['debit_balance'], fm)
        self.sheet.write('H{r}'.format(r=row), data['credit_balance'], fm)

        # Update data for total row
        if v['level'] == 1:
            self.dr_begin_period += data['dr_begin_period']
            self.cr_begin_period += data['cr_begin_period']
            self.dr_in_period += data['dr_in_period']
            self.cr_in_period += data['cr_in_period']
            self.debit_balance += data['debit_balance']
            self.credit_balance += data['credit_balance']

    def get_account_data(self):
        params = self.get_params_for_query()

        sql = """
WITH recursive _tree AS (
    SELECT _p.id AS id
        , _p.name
        , _p.parent_id AS rel_parent
        , 1 AS level
        , _config.group_by_partner
        , array[_p.id] AS _path
        , _p.code
    FROM account_account _p
    LEFT JOIN balance_sheet_config _config
        ON substring(_p.code FOR 3) =_config.code
    WHERE _p.parent_id ISNULL
        AND _p.company_id = {company_id}

    UNION ALL

    SELECT _c.id AS id
        , _c.name
        , _c.parent_id AS rel_parent
        , p.level + 1 AS level
        , p.group_by_partner
        , p._path || _c.id
        , _c.code
    FROM account_account _c
    JOIN _tree p ON _c.parent_id = p.id
    WHERE _c.company_id = {company_id}
)
SELECT row_to_json(final) FROM (
    SELECT
        acc_sum.account_id
        , acc._path
        , acc.code
        , acc."level"
        , acc.name
        , acc.group_by_partner

        , coalesce(sum(
            CASE WHEN acc.group_by_partner ISNULL OR
                acc.group_by_partner = 'f'
        THEN acc_sum.begin_period ELSE 0 END), 0)
         AS begin_period_by_acc

        , coalesce(sum(
            CASE WHEN acc.group_by_partner = 't' AND acc_sum.begin_period > 0
        THEN acc_sum.begin_period ELSE 0 END), 0) AS dr_begin_period

        , coalesce(sum(
            CASE WHEN acc.group_by_partner = 't' AND acc_sum.begin_period < 0
        THEN -acc_sum.begin_period ELSE 0 END), 0) AS cr_begin_period

        , coalesce(sum(acc_sum.dr_in_period), 0) AS dr_in_period_by_acc

        , coalesce(sum(acc_sum.cr_in_period), 0) AS cr_in_period_by_acc

        , coalesce(sum(
            CASE WHEN acc.group_by_partner = 't' AND acc_sum.end_period > 0
            THEN acc_sum.end_period ELSE 0 END), 0) AS dr_end_period

        , coalesce(sum(
            CASE WHEN acc.group_by_partner = 't' AND acc_sum.end_period < 0
        THEN -acc_sum.end_period ELSE 0 END), 0) AS cr_end_period

    FROM _tree acc
    LEFT JOIN
    (
        SELECT ml.account_id, ml.partner_id
            , coalesce(sum(CASE WHEN ml.date < '{date_from}'
            THEN ml.balance END), 0) AS begin_period

            -- , coalesce(sum(CASE WHEN ml.date >= '{date_from}' AND
            -- ml.date <= '{date_to}' THEN ml.balance END), 0) AS in_period

            , coalesce(sum(CASE WHEN ml.date >= '{date_from}' AND
            ml.date <= '{date_to}' THEN ml.debit END), 0) AS dr_in_period

            , coalesce(sum(CASE WHEN ml.date >= '{date_from}' AND
            ml.date <= '{date_to}' THEN ml.credit END), 0) AS cr_in_period

            , coalesce(sum(CASE WHEN ml.date <= '{date_to}'
            THEN ml.balance END), 0) AS end_period

        FROM account_move_line ml JOIN account_move m ON ml.move_id = m.id
        WHERE m.state IN {target_move}
            -- try to limit the records
            AND ml.account_id IN (SELECT id FROM _tree)
        GROUP BY ml.account_id, ml.partner_id

    ) acc_sum ON acc.id = acc_sum.account_id

    -- WHERE
    --     acc.code LIKE '131%'
    --     AND acc.id NOT IN (3310, 4354)

    GROUP BY acc_sum.account_id, acc._path, acc.code, acc."level"
        , acc.name, acc.group_by_partner
    ORDER BY substring(acc.code FOR 3), acc._path
) final
        """.format(company_id=params['company_id'],
                   date_from=params['date_from'],
                   date_to=params['date_to'],
                   target_move=params['target_move'])
        print sql
        self.env.cr.execute(sql)
        return {'data': self.env.cr.fetchall(), 'disp_acc': params['disp_acc']}

    def _re_update_balance(self, res, path_max_level, prev_acc_id):
        max_lev = path_max_level[prev_acc_id]
        to_do = {k: v for k, v in res.items()
                 if v['path'][0] == prev_acc_id}
        while max_lev > 0:
            lead = {k: v for k, v in to_do.items()
                    if v['level'] == max_lev}
            for k_lead, v_lead in lead.items():
                to_update = {
                    k: v for k, v in to_do.items()
                    if set(v['path']) < set(v_lead['path']) and
                    max_lev == v['level'] + 1
                }
                for k in to_update:
                    res[k]['dr_begin_period'] += v_lead['dr_begin_period']
                    res[k]['cr_begin_period'] += v_lead['cr_begin_period']
                    res[k]['begin_period_by_acc'] += \
                        v_lead['begin_period_by_acc']
                    res[k]['dr_in_period_by_acc'] += \
                        v_lead['dr_in_period_by_acc']
                    res[k]['cr_in_period_by_acc'] += \
                        v_lead['cr_in_period_by_acc']
                    res[k]['dr_end_period'] += v_lead['dr_end_period']
                    res[k]['cr_end_period'] += v_lead['cr_end_period']
            max_lev -= 1

    def generate_main_content(self):
        # Reset total balance
        self.dr_begin_period = 0
        self.cr_begin_period = 0
        self.dr_in_period = 0
        self.cr_in_period = 0
        self.debit_balance = 0
        self.credit_balance = 0

        records = self.get_account_data()
        # The records is sorted correctly, therefore we should use OrderedDict
        # to be sure the data will be proceeded in order
        res = OrderedDict()
        path_max_level, prev_acc_id = {}, 0

        for line in records['data']:
            line = line[0]
            path, level = line['_path'], line.get('level', 1)
            if prev_acc_id == 0:
                prev_acc_id = path[0]

            res.update({line['code']:
                        {'name': line['name'], 'level': level,
                         'dr_begin_period': line['dr_begin_period'],
                         'cr_begin_period': line['cr_begin_period'],
                         'begin_period_by_acc': line['begin_period_by_acc'],
                         'dr_in_period_by_acc': line['dr_in_period_by_acc'],
                         'cr_in_period_by_acc': line['cr_in_period_by_acc'],
                         'dr_end_period': line['dr_end_period'],
                         'cr_end_period': line['cr_end_period'],
                         'group_by_partner': line['group_by_partner'],
                         'path': path,
                         }})

            if path[0] not in path_max_level or \
                    level > path_max_level[path[0]]:
                path_max_level.update({path[0]: level})

            # Trigger to update balance
            if path[0] != prev_acc_id:
                self._re_update_balance(res, path_max_level, prev_acc_id)

                prev_acc_id = path[0]

        # Update the last parent account
        self._re_update_balance(res, path_max_level, prev_acc_id)

        # Write data to xlsx file
        disp_acc = records['disp_acc']
        currency = self.object.env.user.company_id.currency_id

        for k, v in res.items():
            dr_in_period = cr_in_period = dr_begin_period = cr_begin_period = 0
            if v['group_by_partner']:
                dr_begin_period = v['dr_begin_period']
                cr_begin_period = v['cr_begin_period']
            else:
                if v['begin_period_by_acc'] > 0:
                    dr_begin_period = v['begin_period_by_acc']
                else:
                    cr_begin_period = abs(v['begin_period_by_acc'])
            dr_in_period = v['dr_in_period_by_acc']
            cr_in_period = v['cr_in_period_by_acc']

            begin_balance = dr_begin_period - cr_begin_period

            debit_balance = dr_begin_period + dr_in_period
            credit_balance = cr_begin_period + cr_in_period
            if v['group_by_partner']:
                debit_balance = v['dr_end_period']
                credit_balance = v['cr_end_period']

            end_balance = debit_balance - credit_balance

            if disp_acc == 'movement':
                if currency.is_zero(cr_in_period) and \
                        currency.is_zero(dr_in_period) and \
                        currency.is_zero(begin_balance):
                    continue
            elif disp_acc == 'not_zero':
                if currency.is_zero(begin_balance) and \
                        currency.is_zero(cr_in_period) and \
                        currency.is_zero(dr_in_period) and \
                        currency.is_zero(end_balance):
                    continue

            if not v['group_by_partner']:
                debit_balance = end_balance if end_balance > 0 else 0
                credit_balance = abs(end_balance) if end_balance < 0 else 0
            self._write_line(k, v, self.ws_row, data={
                             'dr_begin_period': dr_begin_period,
                             'cr_begin_period': cr_begin_period,
                             'dr_in_period': dr_in_period,
                             'cr_in_period': cr_in_period,
                             'debit_balance': debit_balance,
                             'credit_balance': credit_balance,
                             })
            self.ws_row += 1

        # Write total row
        self.sheet.write('A{r}'.format(r=self.ws_row), '', self.format_table)
        self.sheet.write('B{r}'.format(r=self.ws_row),
                         u'Tổng cộng', self.format_table_bold)
        self.sheet.write('C{r}'.format(r=self.ws_row),
                         self.dr_begin_period, self.format_table_bold)
        self.sheet.write('D{r}'.format(r=self.ws_row),
                         self.cr_begin_period, self.format_table_bold)
        self.sheet.write('E{r}'.format(r=self.ws_row),
                         self.dr_in_period, self.format_table_bold)
        self.sheet.write('F{r}'.format(r=self.ws_row),
                         self.cr_in_period, self.format_table_bold)
        self.sheet.write('G{r}'.format(r=self.ws_row),
                         self.debit_balance, self.format_table_bold)
        self.sheet.write('H{r}'.format(r=self.ws_row),
                         self.credit_balance, self.format_table_bold)

        return True

    def generate_footer_content(self):
        self.ws_row += 2

        self.sheet.merge_range(
            'F{r}:H{r}'.format(r=self.ws_row),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.ws_row += 1

        self.sheet.merge_range(
            'A{r}:B{r}'.format(r=self.ws_row),
            u'Người lập biểu', self.format_bold_center)

        self.sheet.merge_range(
            'C{r}:E{r}'.format(r=self.ws_row),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'F{r}:H{r}'.format(r=self.ws_row),
            u'Giám đốc', self.format_bold_center)
        self.ws_row += 1

        self.sheet.merge_range(
            'A{r}:B{r}'.format(r=self.ws_row),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'C{r}:E{r}'.format(r=self.ws_row),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'F{r}:H{r}'.format(r=self.ws_row),
            u'(Ký, họ tên, đóng dấu)', self.format_center)


TrialBalanceReportXlsx(
    'report.report_trial_balance_xlsx',
    'account.report.trial.balance.wizard')
