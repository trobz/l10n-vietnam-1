# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from . import account_purchases_journal_report as APJR
import logging

ACCOUNT_MAPPING = [{'dr': [131], 'cr': [511, 3331]}]
_logger = logging.getLogger(__name__)


class account_sales_journal_xls_parser(
        APJR.account_purchases_journal_xls_parser):

    def __init__(self, cr, uid, name, context):
        super(account_sales_journal_xls_parser, self).__init__(
            cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'is_purchases_journal': False,
        })

    def get_lines_data(self):
        """
        Get all account data
        debit_account is 11*
            credit_account
        """
        account_obj = self.pool.get('account.account')
        date_info = self.get_date()

        # search all accounts
        acc_dr_ids = []
        acc_cr_ids = []
        for account in ACCOUNT_MAPPING:
            for acc in account.get('dr', []):
                acc_dr_ids += account_obj.search(
                    self.cr,
                    self.uid, [('code', '=like', '%s%%' % acc)])
        # get all account id of list account credit
            for acc in account.get('cr', []):
                acc_cr_ids += account_obj.search(
                    self.cr,
                    self.uid, [('code', '=like', '%s%%' % acc)])

        params = {
            'amv_dr_state': self.get_wizard_data()[
                'target_move'] == 'posted' and
            ''' AND amv_dr.state = 'posted' ''' or '',
            'amv_cr_state': self.get_wizard_data()[
                'target_move'] == 'posted' and
            ''' AND amv_cr.state = 'posted' ''' or '',
            'date_from': date_info['date_from_date'],
            'date_to': date_info['date_to_date'],
            'acc_dr_ids': tuple(acc_dr_ids + [-1, -1]),
            'acc_cr_ids': tuple(acc_cr_ids + [-1, -1])
        }

#         SQL = """
#
#         /* Case 1: counterpart_id is set on Cr
#
#             + one-one
#                 Dr    |    Cr
#                 1111  |
#                       |    131    counterpart_id
#         */
#
#         SELECT  move_cr.date_created,
#             amv.name as serial,
#             move_cr.date as effective_date,
#             move_cr.name as description,
#             dr_account.code as debit_code,
#             cr_account.code as credit_code,
#             move_dr.debit as amount
#         FROM
#             account_move_line move_dr
#             LEFT JOIN account_move amv ON amv.id= move_dr.move_id
#             JOIN account_move_line move_cr
#                  ON move_cr.counter_move_id = move_dr.id
#             JOIN account_account dr_account
#                  ON move_dr.account_id = dr_account.id
#             JOIN account_account cr_account
#                  ON move_cr.account_id = cr_account.id
#         WHERE
#             move_cr.credit != 0
#             AND move_dr.debit != 0
#             AND move_dr.date >= '%(date_start)s'
#             AND move_dr.date <= '%(date_end)s'
#             AND amv.inventory_valuation_picking_type = 'sale'
#             %(target_move)s
#         """

        SQL = """

        SELECT
            amv_line_dr.date_created,
            amv_dr.name as serial,
            amv_line_dr.date as effective_date,
            amv_line_dr.name as description,
            dr_account.code as debit_code,
            amv_line_dr.debit as amount,
            array_to_string(array_agg(distinct amv_line_cr.code),',')
            AS credit_code

        FROM account_move_line amv_line_dr

        LEFT JOIN (
            SELECT amv_cr.name,amv_cr.id,acc.code, counter_move_id
            FROM account_move_line amv_line_cr
            JOIN account_account acc ON acc.id=amv_line_cr.account_id
            JOIN account_move amv_cr ON amv_cr.id=amv_line_cr.move_id
            WHERE
                amv_line_cr.account_id in %(acc_cr_ids)s
                AND amv_line_cr.credit <> 0.0
                            AND amv_line_cr.date <= '%(date_to)s' AND
                            amv_line_cr.date >= '%(date_from)s'
                            %(amv_cr_state)s

            ) amv_line_cr ON amv_line_cr.counter_move_id=amv_line_dr.id

        JOIN account_account dr_account
        ON amv_line_dr.account_id = dr_account.id
        JOIN account_move amv_dr
        ON amv_dr.id=amv_line_dr.move_id
        JOIN account_journal acj
        ON acj.id = amv_dr.journal_id

    WHERE
        amv_line_dr.account_id in %(acc_dr_ids)s
        AND amv_line_dr.debit <> 0.0
                AND amv_line_dr.date <= '%(date_to)s' AND
                amv_line_dr.date >= '%(date_from)s'
                %(amv_dr_state)s
                AND acj.type = 'sale'

        GROUP BY
            amv_line_dr.id,amv_line_cr.counter_move_id,
            amv_line_dr.date_created,
            amv_dr.name,
            amv_dr.create_date,
            amv_line_dr.date,
            amv_line_dr.name,
            dr_account.code,
            amv_line_dr.debit

        ORDER BY amv_dr.create_date
        """

#         print "######################### SQL params", SQL % params

        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data


APJR.account_purchases_journal_xls(
    'report.sales_journal_report',
    'account.move',
    parser=account_sales_journal_xls_parser
)
