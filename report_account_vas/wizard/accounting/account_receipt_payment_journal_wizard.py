# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.osv import fields, osv
from openerp.tools.translate import _


class account_receipt_payment_journal_wizard(osv.osv_memory):

    _name = 'receipt.payment.journal.wizard'
    _inherit = "common.ledger"
    _description = 'Print Receipt Report'

    _columns = {
        'type_report': fields.selection(
            [('receipt_journal', 'Receipt Journal'),
             ('payment_journal', 'Payment Journal')],
            'Type',
            required=True
        ),
    }

    _defaults = {
        'type_report': 'receipt_journal',
    }

    def onchange_type(self, cr, uid, ids, type_report, context=None):
        res = {'domain': {}}
        user = self.pool.get('res.users').browse(cr, uid, uid, context=context)
        account_code = "11%"
        sql = """
        SELECT id
        FROM account_account
        WHERE parent_id is not null
        AND type = 'liquidity'
        AND code like '%s'
        AND company_id = %d
        """ % (account_code, user.company_id.id)
        cr.execute(sql)
        chart_account_id = [i[0] for i in cr.fetchall()]
        res['domain'] = {'account': [('id', 'in', chart_account_id)]}
        return res

    def print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}

        data = {
            'model': 'account.move.line',
            'ids': ids,
            'form': self.read(cr, uid, ids[0], context=context),
        }

        report_name = 'account.receipt.journal.report'
        name = 'Receipt Journal Report'
        if data['form']['type_report'] == 'payment_journal':
            report_name = 'account.payment.journal.report'
            name = 'Payment Journal Report'
        if data['form']['filter'] == 'filter_date':
            if data['form']['date_from'] > data['form']['date_to']:
                raise osv.except_osv(_('Warning !'), _(
                    "Start Date must be before End Date !"))
        elif data['form']['filter'] == 'filter_period':
            period_pool = self.pool.get('account.period')
            period_start = period_pool.browse(
                cr, uid, data['form']['period_from'][0])
            period_end = period_pool.browse(
                cr, uid, data['form']['period_to'][0])
            if period_start.date_start > period_end.date_stop:
                raise osv.except_osv(_('Warning !'), _(
                    "Start Period must be before End Period !"))

        return {
            'type': 'ir.actions.report.xml',
            'report_name': report_name,
            'datas': data,
            'name': name
        }


account_receipt_payment_journal_wizard()
