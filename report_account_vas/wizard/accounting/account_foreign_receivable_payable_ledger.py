# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import fields, osv


class account_foreign_receivable_payable_ledger(osv.osv_memory):
    _inherit = 'account.detailed.payable.receivable.balance'
    _name = 'account.foreign.receivable.payable.ledger'
    _description = 'Print Foreign Receivable/Payable Ledger'

    _columns = {
        'partner_id': fields.many2one('res.partner', 'Partner'),
        'foreign_currency_id': fields.many2one(
            'res.currency',
            'Currency',
            required=True
        ),
        'journal_ids': fields.many2many(
            'account.journal',
            'account_foreign_receivable_payable_ledger_journal_rel',
            'account_id',
            'journal_id',
            'Journals',
            required=True
        ),
    }

    def onchange_result_selection(
            self, cr, uid, ids, result_selection, context=None):
        if context is None:
            context = {}
        res = {}
        if result_selection:
            if result_selection == 'customer':
                res.update({
                    'value': {'partner_id': False},
                    'domain': {'partner_id': [('customer', '=', True)]}
                })
            else:
                res.update({
                    'value': {'partner_id': False},
                    'domain': {'partner_id': [('supplier', '=', True)]}
                })
        return res

    def pre_print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}

        data['form'].update(self.read(
            cr, uid, ids, [
                'result_selection', 'partner_id',
                'foreign_currency_id', 'fiscalyear_id', 'period_from',
                'period_to'], context=context)[0])
        return data

    def _print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}
        data = self.pre_print_report(cr, uid, ids, data, context=context)
        report_name = 'foreign_receivable_payable_ledger_report'
        name = 'Foreign Receivable Ledger Report'
        if data['form']['result_selection'] == 'supplier':
            name = 'Foreign Payable Ledger Report'

        return {
            'type': 'ir.actions.report.xml',
            'report_name': report_name,
            'datas': data,
            'name': name
        }


account_foreign_receivable_payable_ledger()
