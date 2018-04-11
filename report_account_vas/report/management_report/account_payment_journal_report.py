# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from . import account_receipt_journal_report as ARJR
import logging
_logger = logging.getLogger(__name__)


class account_payment_journal_xls_parser(
        ARJR.account_receipt_journal_xls_parser):

    def __init__(self, cr, uid, name, context):
        super(account_payment_journal_xls_parser, self).__init__(
            cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'is_receipt_journal': False
        })

    def get_debit_account_data(self):
        """
        Get all account data
        debit_account is 11*
            credit_account
        """
        date_info = self.get_date()
        account_id = self.get_wizard_data()['account'] or -1
        params = {'target_move': '', 'account_id': account_id,
                  'date_start': date_info['date_from_date'],
                  'date_end': date_info['date_to_date']}

        SQL = """
        /* Case 1: counterpart_id is set on Dr
            + one-many
                  Dr    |    Cr
                        |    1111
                  3311  |            counterpart_id
                  3311  |            counterpart_id

        */

        SELECT  move_dr.date_created,
        amv.name as serial,
        move_dr.date as effective_date,
        move_dr.name as description,
        account_account.code as code,
        move_dr.debit as amount
        FROM
            account_move_line move_cr
        LEFT JOIN account_move amv ON amv.id= move_cr.move_id
        JOIN account_move_line move_dr ON move_dr.counter_move_id = move_cr.id
        JOIN account_account ON move_dr.account_id = account_account.id
        WHERE
        move_dr.debit != 0
        AND move_cr.account_id = %(account_id)s
        AND move_cr.date >= '%(date_start)s' AND move_cr.date <= '%(date_end)s'
        %(target_move)s

    UNION

        /* Case 2: counterpart_id is set on Cr
        + one2many
            Dr      |    Cr
                    |    1111        counterpart_id
            3311    |
                    |    1121        counterpart_id


        + one-one
                Dr     |    Cr
                       |    1111    counterpart_id
               3311    |
        */

        SELECT  move_cr.date_created,
        amv.name as serial,
        move_cr.date as effective_date,
        move_cr.name as description,
        account_account.code as code,
        move_cr.credit as amount
        FROM
            account_move_line move_cr
        LEFT JOIN account_move amv ON amv.id= move_cr.move_id
        JOIN account_move_line move_dr ON move_cr.counter_move_id = move_dr.id
        JOIN account_account ON move_dr.account_id = account_account.id
        WHERE
        move_dr.debit != 0
        AND move_cr.account_id = %(account_id)s
        AND move_cr.date >= '%(date_start)s' AND move_cr.date <= '%(date_end)s'
        %(target_move)s

        """

        if self.get_wizard_data()['target_move'] == 'posted':
            params['target_move'] = "AND amv.state = 'posted'"

        # order by date_created
        SQL = SQL + " ORDER BY date_created"

#         print ">>> SQL ============", SQL % params
        self.cr.execute(SQL % params)
        data = self.cr.dictfetchall()
        return data


class account_payment_journal_xls(ARJR.account_receipt_journal_xls):

    def __init__(
            self, name, table, rml=False,
            parser=False, header=True, store=False):
        super(account_payment_journal_xls, self).__init__(
            name, table, rml, parser, header, store)


account_payment_journal_xls(
    'report.account.payment.journal.report',
    'account.move.line',
    parser=account_payment_journal_xls_parser
)
