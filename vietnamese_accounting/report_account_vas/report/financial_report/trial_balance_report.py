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


import time
from openerp.addons.report_base_vn.report import report_base_vn
from datetime import datetime


class Parser(report_base_vn.Parser):
    def __init__(self, cr, uid, name, context=None):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'trial_balance_report'
        self.debit_beginning_of_period = 0.0
        self.credit_beginning_of_period = 0.0
        self.debit_in_period = 0.0
        self.credit_in_period = 0.0
        self.debit_balance = 0.0
        self.credit_balance = 0.0

        self.localcontext.update({
            'time': time,
            'lines': self.lines,
            'get_date': self._get_date,
            'get_company': self._get_company,
            'get_total': self._get_total,
        })
        self.context = context

    def _get_total(self, type='', debit=True):  # @UnresolvedImport #@ReservedAssignment
        """
        @param: string type in ['begin', 'in', 'end']
                debit = True or False
        """

        if type == 'begin':
            if debit:
                return self.debit_beginning_of_period
            else:
                return self.credit_beginning_of_period
        if type == 'in':
            if debit:
                return self.debit_in_period
            else:
                return self.credit_in_period
        if type == 'end':
            if debit:
                return self.debit_balance
            else:
                return self.credit_balance
        return 0.0

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
        return {'date_from': datetime.strptime(res['date_from'], '%Y-%m-%d').strftime('%d-%m-%Y'),
                'date_to': datetime.strptime(res['date_to'], '%Y-%m-%d').strftime('%d-%m-%Y')}

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
        return super(Parser, self).set_context(objects, data, new_ids, report_type=report_type)

    def _get_company(self):
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        address_list = [res.company_id.street or '',
                        res.company_id.street2 or '',
                        res.company_id.city or '',
                        res.company_id.state_id and res.company_id.state_id.name or '',
                        res.company_id.country_id and res.company_id.country_id.name or '',
                        ]
        while '' in address_list:
            address_list.remove('')
        address = ', '.join(address_list)
        currency = res.company_id.currency_id.name
        return {'name': name, 'address': address, 'currency': currency}

    def lines(self, form, ids=None, done=None):
        """
        Get debit, credit for every account
        result:
            [{'id': account_id, 'code', 'name', 'type', 'debit_beginning_of_period', 'credit_beginning_of_period',
            'debit_in_period', 'credit_in_period', 'debit_balance', 'credit_balance'}]
        """
        result_acc = []
        obj_account = self.pool.get('account.account')
        obj_fiscalyear = self.pool.get('account.fiscalyear').browse(
            self.cr, self.uid, form['fiscalyear_id'])
        if not ids:
            ids = self.ids
        if not ids:
            return []
        if not done:
            done = {}
        ctx = self.context.copy()
        ctx['fiscalyear'] = form['fiscalyear_id']
        if form['filter'] == 'filter_period':
            period_pool = self.pool.get('account.period')
            period_start = period_pool.browse(
                self.cr, self.uid, form['period_from'])
            period_end = period_pool.browse(
                self.cr, self.uid, form['period_to'])
            ctx['date_from'] = period_start.date_start
            ctx['date_to'] = period_end.date_stop
        elif form['filter'] == 'filter_date':
            ctx['date_from'] = form['date_from']
            ctx['date_to'] = form['date_to']
        else:
            ctx['date_from'] = obj_fiscalyear.date_start
            ctx['date_to'] = obj_fiscalyear.date_stop
        ctx['state'] = form['target_move']
        self.date_from = ctx['date_from']
        self.date_to = ctx['date_to']
        disp_acc = form['display_account']
        currency_obj = self.pool.get('res.currency')
        state = form['target_move'] == 'posted' and "and m.state = 'posted'" or ''
        parents = ids
        for parent in parents:
            acc_id = self.pool.get('account.account').browse(
                self.cr, self.uid, parent)
            currency = acc_id.currency_id and acc_id.currency_id or acc_id.company_id.currency_id
            child_ids = obj_account._get_children_and_consol(
                self.cr, self.uid, [parent], ctx)
            if child_ids:
                sql = '''
                    SELECT acc.id, acc.code, acc.name, acc_sum.debitbeginningofperiod, acc_sum.creditbeginningofperiod,
                            acc_sum.debitinperiod, acc_sum.creditinperiod, acc.type
                    FROM account_account acc LEFT JOIN (
                                SELECT ml.account_id,
                                    COALESCE(SUM(CASE WHEN ml.date < '%s' THEN ml.debit END),0) AS debitbeginningofperiod,
                                    COALESCE(SUM(CASE WHEN ml.date < '%s' THEN ml.credit END),0) AS creditbeginningofperiod,
                                    COALESCE(SUM(CASE WHEN ml.date >= '%s' AND ml.date <= '%s' THEN ml.debit END),0) AS debitinperiod,
                                    COALESCE(SUM(CASE WHEN ml.date >= '%s' AND ml.date <= '%s' THEN ml.credit END),0) AS creditinperiod
                                FROM account_move_line ml JOIN account_move m ON ml.move_id = m.id
                                WHERE ml.account_id IN (%s) %s
                                GROUP BY ml.account_id ) acc_sum
                        ON acc.id = acc_sum.account_id
                    WHERE acc.level > 1
                    ORDER BY acc.code;

                ''' % (ctx['date_from'], ctx['date_from'],
                       ctx['date_from'], ctx['date_to'],
                       ctx['date_from'], ctx['date_to'],
                       "," .join(map(str, child_ids)),
                       state)
                self.cr.execute(sql)
                account_list = self.cr.fetchall()
                for account in account_list:
                    if account[1][0] == '0':
                        continue
                    res = {
                        'id': account[0],
                        'code': account[1],
                        'name': account[2],
                        'type': account[7],
                        'debit_beginning_of_period': account[3] or 0.0,
                        'credit_beginning_of_period': account[4] or 0.0,
                        'debit_in_period': account[5] or 0.0,
                        'credit_in_period': account[6] or 0.0,
                        'debit_balance': 0.0,
                        'credit_balance': 0.0,
                    }
                    # if account have children recompute debit and credit
                    if account[7] == 'view':
                        child_ids = obj_account._get_children_and_consol(
                            self.cr, self.uid, [account[0]], ctx)
                        if child_ids:
                            for child_account in account_list:
                                if child_account[0] in child_ids and child_account[0] != account[0]:
                                    res['debit_beginning_of_period'] += child_account[3] or 0.0
                                    res['credit_beginning_of_period'] += child_account[4] or 0.0
                                    res['debit_in_period'] += child_account[5] or 0.0
                                    res['credit_in_period'] += child_account[6] or 0.0
                    # compute balance amount for every account
                    begin_balance = res['debit_beginning_of_period'] - \
                        res['credit_beginning_of_period']
                    end_balance = begin_balance + \
                        res['debit_in_period'] - res['credit_in_period']
                    if '1' == res['code'][0] or '2' == res['code'][0]:
                        if end_balance >= 0:
                            res['debit_balance'] = end_balance
                        else:
                            res['credit_balance'] = -1.0 * end_balance
                        if begin_balance >= 0:
                            res['debit_beginning_of_period'] = begin_balance
                            res['credit_beginning_of_period'] = 0.0
                        else:
                            res['debit_beginning_of_period'] = 0.0
                            res['credit_beginning_of_period'] = - \
                                1.0 * begin_balance
                    else:
                        if end_balance < 0:
                            res['credit_balance'] = -1.0 * end_balance
                            res['debit_balance'] = 0.0
                        else:
                            res['debit_balance'] = end_balance
                            res['credit_balance'] = 0.0
                        if begin_balance < 0:
                            res['credit_beginning_of_period'] = - \
                                1.0 * begin_balance
                            res['debit_beginning_of_period'] = 0.0
                        else:
                            res['debit_beginning_of_period'] = begin_balance
                            res['credit_beginning_of_period'] = 0.0

                    if disp_acc == 'movement':
                        if not currency_obj.is_zero(self.cr, self.uid, currency, res['credit_in_period']) or \
                                not currency_obj.is_zero(self.cr, self.uid, currency, res['debit_in_period']) or \
                                not currency_obj.is_zero(self.cr, self.uid, currency, begin_balance):
                            result_acc.append(res)
                    elif disp_acc == 'not_zero':
                        if not currency_obj.is_zero(self.cr, self.uid, currency, begin_balance) or \
                                not currency_obj.is_zero(self.cr, self.uid, currency, res['credit_in_period']) or \
                                not currency_obj.is_zero(self.cr, self.uid, currency, res['debit_in_period']) or \
                                not currency_obj.is_zero(self.cr, self.uid, currency, res['debit_balance'] - res['credit_balance']):
                            result_acc.append(res)
                    else:
                        result_acc.append(res)
            # compute total in report
            list_child_ids = []
            for account in result_acc:
                child_ids = obj_account._get_children_and_consol(
                    self.cr, self.uid, [account['id']], ctx)
                child_ids.remove(account['id'])
                list_child_ids += child_ids
                if account['id'] not in list_child_ids:
                    self.debit_beginning_of_period += account['debit_beginning_of_period']
                    self.credit_beginning_of_period += account['credit_beginning_of_period']
                    self.debit_in_period += account['debit_in_period']
                    self.credit_in_period += account['credit_in_period']
                    self.debit_balance += account['debit_balance']
                    self.credit_balance += account['credit_balance']

        return result_acc

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
