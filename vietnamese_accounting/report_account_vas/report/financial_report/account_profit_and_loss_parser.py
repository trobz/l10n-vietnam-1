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
from openerp.addons.report_xls.utils import _render  # @UnresolvedImport
from .. import report_xls_utils
import time
from openerp.addons.report_base_vn.report import report_base_vn
from datetime import datetime
from datetime import timedelta
# import ast
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT


class Parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context=None):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'account_profit_and_loss_report_xls'
        self.result = {}
        self.localcontext.update({
            'time': time,
            'get_date': self._get_date,
            'get_result': self.get_result,
            'get_company': self._get_company,
        })
        self.context = context

    def _get_date(self, data):
        res = {}
        obj_fiscalyear = self.pool.get('account.fiscalyear').browse(
            self.cr, self.uid, data['form']['fiscalyear_id'])
        if data['form']['filter'] == 'filter_period':
            period_pool = self.pool.get('account.period')
            period_start = period_pool.browse(
                self.cr, self.uid, data['form']['period_from'])
            period_end = period_pool.browse(
                self.cr, self.uid, data['form']['period_to'])
            res.update({
                'date_from': period_start.date_start,
                'date_to': period_end.date_stop,
            })
        elif data['form']['filter'] == 'filter_date':
            res.update({
                'date_from': data['form']['date_from'],
                'date_to': data['form']['date_to'],
            })
        else:
            res.update({
                'date_from': obj_fiscalyear.date_start,
                'date_to': obj_fiscalyear.date_stop,
            })
        if not res.get('date_from', False) or not res.get('date_to', False):
            return {'date_from': '', 'date_to': ''}

        self.date_from_date = res['date_from']
        self.last_date_from_date = datetime.strptime(
            res['date_from'], DEFAULT_SERVER_DATE_FORMAT) - timedelta(days=365)
        self.date_to_date = res['date_to']
        self.last_date_to_date = datetime.strptime(
            res['date_to'], DEFAULT_SERVER_DATE_FORMAT) - timedelta(days=365)

        return {

            'date_from': datetime.strptime(res['date_from'], DEFAULT_SERVER_DATE_FORMAT).strftime('%d-%m-%Y'),
            'date_to': datetime.strptime(res['date_to'], DEFAULT_SERVER_DATE_FORMAT).strftime('%d-%m-%Y'),
            'last_date_from': self.last_date_from_date.strftime('%d-%m-%Y'),
            'last_date_to': self.last_date_to_date.strftime('%d-%m-%Y')
        }

    def _get_start_date(self, data):
        if data.get('form', False) and data['form'].get('date_from', False):
            return data['form']['date_from']

        return ''

    def _get_end_date(self, data):
        if data.get('form', False) and data['form'].get('date_to', False):
            return data['form']['date_to']
        return ''

    def set_context(self, objects, data, ids, report_type=None):
        new_ids = ids
        if (data['model'] == 'ir.ui.menu'):
            new_ids = 'chart_account_id' in data['form'] and [
                data['form']['chart_account_id']] or []
            objects = self.pool.get('account.account').browse(
                self.cr, self.uid, new_ids)
        # set date
        self._get_date(data)
        # compute result of every line
        self.compute_data(data)

        return super(Parser, self).set_context(objects, data, new_ids, report_type=report_type)

    def _get_company(self):
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        address_list = [res.company_id.street or '',
                        res.company_id.street2 or '', res.company_id.city or '']
        while '' in address_list:
            address_list.remove('')
        address = ', '.join(address_list)
        currency = res.company_id.currency_id.name
        return {'name': name, 'address': address, 'currency': currency}

    def get_account_data(self, acc_dr_ids, acc_cr_ids, target_move=False):

        if not acc_dr_ids or not acc_cr_ids:
            return 0.0
        # in case of account move line have counter_move_id is null
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        #    cr 33311: 1    counterpart_id
        #    cr 511: 14     counterpart_id
        # =====> sum = credit cr (33311,511)
        params = {'acc_dr_ids': tuple(acc_dr_ids + [-1, -1]), 'acc_cr_ids': tuple(acc_cr_ids + [-1, -1]),
                  'date_from': self.date_from_date, 'date_to': self.date_to_date,
                  'last_date_from': self.last_date_from_date, 'last_date_to': self.last_date_to_date,
                  'amv_cr_state': target_move == 'posted' and ''' AND amv_cr.state = 'posted' ''' or '',
                  'amv_dr_state': target_move == 'posted' and ''' AND amv_dr.state = 'posted' ''' or ''}
        sql = '''
            SELECT COALESCE(SUM(CASE WHEN move_cr.date <= '%(date_to)s' AND move_cr.date >= '%(date_from)s' THEN move_cr.credit END),0) as amount,
                   COALESCE(SUM(CASE WHEN move_cr.date <= '%(last_date_to)s' AND move_cr.date >= '%(last_date_from)s' THEN move_cr.credit END),0) as last_amount

            FROM account_move_line move_cr
            LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
            WHERE move_cr.account_id in %(acc_cr_ids)s
                    AND move_cr.credit <> 0.0
                    AND ((move_cr.date <= '%(date_to)s' AND move_cr.date >= '%(date_from)s') OR
                    (move_cr.date <= '%(last_date_to)s' AND move_cr.date >= '%(last_date_from)s'))
                    %(amv_cr_state)s
                    AND move_cr.counter_move_id in (
                        SELECT move_dr.id
                        FROM account_move_line move_dr
                        LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
                        WHERE move_dr.account_id in %(acc_dr_ids)s
                                AND move_dr.counter_move_id is null
                                AND  move_dr.debit <> 0.0
                                AND ((move_dr.date <= '%(date_to)s' AND move_dr.date >= '%(date_from)s') OR
                                    (move_dr.date <= '%(last_date_to)s' AND move_dr.date >= '%(last_date_from)s'))
                                 %(amv_dr_state)s
                )
                '''

        # in case of account move line have counter_move_id
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        # dr 515: 1
        #    cr 33311: 16
        # =====> sum = dr debit (111,511)
        sql += '''
            UNION ALL

            SELECT COALESCE(SUM(CASE WHEN move_dr.date <= '%(date_to)s' AND move_dr.date >= '%(date_from)s' THEN move_dr.debit END),0) as amount,
                    COALESCE(SUM(CASE WHEN move_dr.date <= '%(last_date_to)s' AND move_dr.date >= '%(last_date_from)s' THEN move_dr.debit END),0) as last_amount
            FROM account_move_line move_dr
            LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
            WHERE move_dr.account_id in %(acc_dr_ids)s
                    AND move_dr.debit <> 0.0
                    AND ((move_dr.date <= '%(date_to)s' AND move_dr.date >= '%(date_from)s') OR
                        (move_dr.date <= '%(last_date_to)s' AND move_dr.date >= '%(last_date_from)s'))
                     %(amv_dr_state)s

                    AND move_dr.counter_move_id in (
                        SELECT move_cr.id
                        FROM account_move_line move_cr
                        LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
                        WHERE move_cr.account_id in %(acc_cr_ids)s
                                AND move_cr.counter_move_id is null
                                AND  move_cr.credit <> 0.0
                                AND ((move_cr.date <= '%(date_to)s' AND move_cr.date >= '%(date_from)s') OR
                                    (move_cr.date <= '%(last_date_to)s' AND move_cr.date >= '%(last_date_from)s'))
                         %(amv_cr_state)s
                    )'''

        sql = sql % params

        self.cr.execute(sql)
        result = self.cr.fetchall()
        current_values = sum([item[0] for item in result if item[0]])
        last_values = sum([item[1] for item in result if item[1]])
        return current_values, last_values

    def compute_data(self, data):
        account_obj = self.pool.get('account.account')
        profit_and_loss_config_pool = self.pool['profit.and.loss.config']
        profit_loss_ids = profit_and_loss_config_pool.search(
            self.cr, self.uid, [])
        profit_loss_objs = profit_and_loss_config_pool.browse(
            self.cr, self.uid, profit_loss_ids)

        for profit_loss_obj in profit_loss_objs:
            # for each element of profit and loss
            key = int(profit_loss_obj.code)
            account_ids = [acc.id for acc in profit_loss_obj.account_ids]
            counterpart_account_ids = [
                acc.id for acc in profit_loss_obj.counterpart_account_ids]

            # in case account ids is null, it means getting all account_ids
            all_account_ids = account_obj.search(self.cr, self.uid, [])
            account_ids = account_ids or all_account_ids
            counterpart_account_ids = counterpart_account_ids or all_account_ids

            # default, is_debit=True : get all total debit

            debit_account_ids = account_ids
            credit_account_ids = counterpart_account_ids

            # in case is_credit=True : get all total credit
            if profit_loss_obj.is_credit:
                debit_account_ids = counterpart_account_ids
                credit_account_ids = account_ids

            target_move = data['form'][
                'target_move'] == 'posted' and 'posted' or False
            current_values, last_values = self.get_account_data(
                debit_account_ids, credit_account_ids, target_move)

            # in case is_exception: get total_debit - total_credit
            if profit_loss_obj.exception:
                inverted_current_values, inverted_last_values = self.get_account_data(
                    credit_account_ids, debit_account_ids, target_move)
                # total_debit - total_credit
                current_values -= inverted_current_values
                last_values -= inverted_last_values
            self.result.update(
                {(key, 'now'): current_values, (key, 'last'): last_values})

        return True

    def get_result(self, key, state):
        """
        @param: key: is Ma so in template report
                state = 'now' or 'last': get amount in period now
        """
        return self.result.get((key, state), 0.0)


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
