# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.addons.report_base_vn.report import report_base_vn
from datetime import datetime, date
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT


class Parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'cash_flow_indirect_report'
        self.wizard_data = {}
        self.result = {}
        self.date_from_now = ''
        self.date_to_now = ''
        self.date_from_last = ''
        self.date_to_last = ''
        self.account_data = {}
        self.localcontext.update({
            'get_date_value': self.get_date_value,
            'get_lines': self.get_lines,
        })

    def set_context(self, objects, datas, ids, report_type=None):
        self.wizard_data = {}
        if datas:
            self.wizard_data['fiscalyear_id'] = 'fiscalyear_id' in datas[
                'form'] and datas['form']['fiscalyear_id'][0] or False
            self.wizard_data['chart_account_id'] = 'chart_account_id' in datas[
                'form'] and datas['form']['chart_account_id'][0] or False
            self.wizard_data['target_move'] = 'target_move' in datas[
                'form'] and datas['form']['target_move'] or ''
            self.wizard_data['type_report'] = 'type_report' in datas[
                'form'] and datas['form']['type_report'] or False
            self.wizard_data['account'] = 'account' in datas[
                'form'] and datas['form']['account'] or False
            self.wizard_data['filter'] = 'filter' in datas[
                'form'] and datas['form']['filter'] or False
            if datas['form']['filter'] == 'filter_date':
                self.wizard_data['date_from'] = datas['form']['date_from']
                self.wizard_data['date_to'] = datas['form']['date_to']
            elif datas['form']['filter'] == 'filter_period':
                self.wizard_data['period_from'] = datas[
                    'form']['period_from'][0]
                self.wizard_data['period_to'] = datas['form']['period_to'][0]
            self._get_date()
            # compute all line in report
            self.compute_result()
        return super(Parser, self).set_context(
            objects, datas, ids, report_type=report_type)

    def get_date_value(self, key):
        """
        from key, get value of selt.key
        """
        mapping = {'date_from': self.date_from_now.strftime('%d-%m-%Y'),
                   'date_to': self.date_to_now.strftime('%d-%m-%Y')}
        return mapping.get(key, '')

    def _get_date(self):
        if self.wizard_data['filter'] == 'filter_period':
            period_obj = self.pool.get('account.period')
            period_start = period_obj.browse(
                self.cr, self.uid, self.wizard_data['period_from'])
            period_end = period_obj.browse(
                self.cr, self.uid, self.wizard_data['period_to'])
            self.date_from_now = period_start.date_start
            self.date_to_now = period_end.date_stop
        elif self.wizard_data['filter'] == 'filter_date':
            self.date_from_now = self.wizard_data['date_from']
            self.date_to_now = self.wizard_data['date_to']
        else:
            fiscalyear = self.pool.get('account.fiscalyear').browse(
                self.cr, self.uid, self.wizard_data['fiscalyear_id'])
            self.date_from_now = fiscalyear.date_start
            self.date_to_now = fiscalyear.date_stop
        # get last year
        self.date_from_now = datetime.strptime(
            self.date_from_now, DEFAULT_SERVER_DATE_FORMAT)
        self.date_to_now = datetime.strptime(
            self.date_to_now, DEFAULT_SERVER_DATE_FORMAT)
        self.date_from_last = date(
            self.date_from_now.year - 1,
            self.date_from_now.month, self.date_from_now.day)
        self.date_to_last = date(
            self.date_to_now.year - 1,
            self.date_to_now.month, self.date_to_now.day)
        return True

    def get_total_debit_credit_data(
            self, acc_dr_ids, acc_cr_ids, target_move=False):
        """
        compute data of total debit or total credit by account_debit accounts
        and related account_credit accounts

        @param acc_dr_ids: account dr ids
        @param acc_cr_ids: account_cr_ids
        @param target_move: target_move of account_entry

        """

        if not acc_dr_ids or not acc_cr_ids:
            return 0.0
        # in case of account move line have counter_move_id is null
        # ex: 1: [{'dr': ('or', ['111']), 'cr': ('or', ['33311'])]
        # dr 111: 15
        #    cr 33311: 1    counterpart_id
        #    cr 511: 14     counterpart_id
        # =====> sum = credit cr (33311,511)
        params = {
            'acc_dr_ids': tuple(acc_dr_ids + [-1, -1]),
            'acc_cr_ids': tuple(acc_cr_ids + [-1, -1]),
            'date_from': self.date_from_now,
            'date_to': self.date_to_now,
            'last_date_from': self.date_from_last,
            'last_date_to': self.date_to_last,
            'amv_cr_state': target_move == 'posted' and
            ''' AND amv_cr.state = 'posted' ''' or '',
            'amv_dr_state': target_move == 'posted' and
            ''' AND amv_dr.state = 'posted' ''' or ''}

        sql = '''
            SELECT
                COALESCE(SUM(
                    CASE WHEN move_cr.date <= '%(date_to)s' AND
                        move_cr.date >= '%(date_from)s'
                        THEN move_cr.credit END),0)
                    AS amount,
                COALESCE(SUM(
                    CASE WHEN move_cr.date <= '%(last_date_to)s' AND
                        move_cr.date >= '%(last_date_from)s'
                        THEN move_cr.credit END),0)
                    AS last_amount

            FROM account_move_line move_cr
            LEFT JOIN account_move amv_cr ON amv_cr.id= move_cr.move_id
            WHERE move_cr.account_id in %(acc_cr_ids)s
                    AND move_cr.credit <> 0.0
                    AND (
                            (
                                move_cr.date <= '%(date_to)s' AND
                                move_cr.date >= '%(date_from)s'
                            ) OR (
                                move_cr.date <= '%(last_date_to)s' AND
                                move_cr.date >= '%(last_date_from)s'
                            )
                        )
                    %(amv_cr_state)s
                    AND move_cr.counter_move_id in (
                        SELECT move_dr.id
                        FROM account_move_line move_dr
                        LEFT JOIN account_move amv_dr
                            ON amv_dr.id= move_dr.move_id
                        WHERE move_dr.account_id in %(acc_dr_ids)s
                                AND move_dr.counter_move_id is null
                                AND  move_dr.debit <> 0.0
                                AND ((
                                        move_dr.date <= '%(date_to)s' AND
                                        move_dr.date >= '%(date_from)s'
                                    ) OR (
                                        move_dr.date <= '%(last_date_to)s'
                                    AND move_dr.date >= '%(last_date_from)s'
                                    ))
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

            SELECT
                COALESCE(SUM(
                    CASE WHEN move_dr.date <= '%(date_to)s' AND
                        move_dr.date >= '%(date_from)s'
                        THEN move_dr.debit END),0)
                    AS amount,
                COALESCE(SUM(
                    CASE WHEN move_dr.date <= '%(last_date_to)s' AND
                        move_dr.date >= '%(last_date_from)s'
                        THEN move_dr.debit END),0)
                    AS last_amount
            FROM account_move_line move_dr
            LEFT JOIN account_move amv_dr ON amv_dr.id= move_dr.move_id
            WHERE move_dr.account_id in %(acc_dr_ids)s
                    AND move_dr.debit <> 0.0
                    AND (
                            (
                                move_dr.date <= '%(date_to)s' AND
                                move_dr.date >= '%(date_from)s'
                            ) OR (
                                move_dr.date <= '%(last_date_to)s' AND
                                move_dr.date >= '%(last_date_from)s'
                            )
                        )
                     %(amv_dr_state)s

                    AND move_dr.counter_move_id in (
                        SELECT move_cr.id
                        FROM account_move_line move_cr
                        LEFT JOIN account_move amv_cr
                            ON amv_cr.id= move_cr.move_id
                        WHERE move_cr.account_id in %(acc_cr_ids)s
                                AND move_cr.counter_move_id is null
                                AND  move_cr.credit <> 0.0
                                AND ((move_cr.date <= '%(date_to)s' AND
                                        move_cr.date >= '%(date_from)s') OR
                                    (move_cr.date <= '%(last_date_to)s' AND
                                    move_cr.date >= '%(last_date_from)s'))
                         %(amv_cr_state)s
                    )'''

        sql = sql % params

        self.cr.execute(sql)
        result = self.cr.fetchall()
        current_values = sum([item[0] for item in result if item[0]])
        last_values = sum([item[1] for item in result if item[1]])
        return current_values, last_values

    def get_total_balance_data(self, acc_ids):
        """
        Compute total balance for an account at the ending period of time
        @param acc_ids: account ids
        """

        result_now = result_last = 0
        for account_id in acc_ids:
            balance_now, balanace_last = self.account_data.get(
                account_id, (0, 0))
            result_now += balance_now
            result_last += balanace_last
        return result_now, result_last

    def compute_total_balance_data_for_all_account(self, target_move):
        """
        Compute total balance for an account at the ending period of time
        This function only compute for accounts which was defined when
            formular_active is True
        """

        state = target_move == 'posted' and (
            " JOIN account_move m ON ml.move_id = m.id "
            "WHERE m.state = 'posted' " or ' WHERE True '
        ) or ''

        params = {
            'date_to_now': self.date_to_now,
            'date_from_now': self.date_from_now,
            'date_to_last': self.date_to_last,
            'date_from_last': self.date_from_last,
            'state': state,
            'account_ids': (-1, -1)
        }

        # we only need accounts which start with [0xx,1xx,2xx,3xx,4xx]
        account_pool = self.pool.get('account.account')
        account_ids = []
        account_types = [0, 1, 2, 3, 4]
        # TODO: really, we only take into which
        for acc_type in account_types:
            account_ids += account_pool.search(
                self.cr, self.uid,
                [('code', '=like', '%s%%' % acc_type)])
        params['account_ids'] = tuple(account_ids) or (-1, -1)

        sql = '''
            SELECT
                ml.account_id,
                COALESCE(SUM(
                    CASE WHEN ml.date >= '%(date_from_now)s' AND
                        ml.date <= '%(date_to_now)s'
                        THEN (ml.debit - ml.credit) END),0)
                    AS balance_now,
                COALESCE(SUM(
                    CASE WHEN ml.date >= '%(date_from_last)s' AND
                        ml.date <= '%(date_to_last)s'
                        THEN (ml.debit - ml.credit) END),0)
                    AS balance_last
            FROM account_move_line ml
             %(state)s
            AND ml.account_id in %(account_ids)s
            GROUP BY ml.account_id;
        ''' % params
        self.cr.execute(sql)
        for line in self.cr.fetchall():
            self.account_data.update({line[0]: (line[1], line[2])})

    def compute_result(self):
        """
        compute all account
        """
        indirect_cash_flow_config_pool = self.pool['indirect.cash.flow.config']
        line_cashflow_config_ids = indirect_cash_flow_config_pool.search(
            self.cr, self.uid, [('parent_id', '=', False)], order="code")
        line_cashflow_config_objs = indirect_cash_flow_config_pool.browse(
            self.cr, self.uid, line_cashflow_config_ids)

        # first, need to compute total balance for some accounts
        target_move = self.wizard_data['target_move']
        self.compute_total_balance_data_for_all_account(target_move)

        # compute value for every item in config parameters
        for line_cashflow_config_obj in line_cashflow_config_objs:

            code = str(line_cashflow_config_obj.code)
            now_total_balance = last_total_balance = 0
            # foreach one child objects
            for config_child in line_cashflow_config_obj.child_ids:

                # if active formula is False
                if not config_child.is_formula_active:
                    # TODO: get debit/credit, invert result if
                    # is_inverted_result is checked
                    dr_account_ids = [
                        account.id for account in config_child.dr_account_ids]
                    cr_account_ids = [
                        account.id for account in config_child.cr_account_ids]
                    now_balance, last_balance = \
                        self.get_total_debit_credit_data(
                            dr_account_ids, cr_account_ids, target_move
                        )
                    if config_child.is_inverted_result:
                        now_balance *= -1
                        last_balance *= -1

                # if active formula = True
                else:
                    # active formula: result = SDCK - SDDK
                    # TODO: get balance
                    to_get_total_balance_account_ids = [
                        account.id for account in config_child.account_ids]
                    now_balance, last_balance = self.get_total_balance_data(
                        to_get_total_balance_account_ids)
                    if config_child.is_positive_difference:
                        now_balance *= -1
                        last_balance *= -1

                # sum value
                now_total_balance += now_balance
                last_total_balance += last_balance

            self.result.update({
                (code, 'last'): last_total_balance,
                (code, 'now'): now_total_balance
            })
        return self.result

    def get_lines(self, key, year):
        """
        @param: key: int ---> column Mã số in template report
        """
        result = self.result.get((str(key), year), 0.0)
        return result
