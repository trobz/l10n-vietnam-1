# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import osv, fields
from openerp.tools.translate import _


class sales_purchases_journal_wizard(osv.osv_memory):

    _name = 'sales.purchases.journal.wizard'
    _inherit = "common.ledger"
    _description = 'Purchases and Sales Journal Wizard'

    _columns = {
        'account': fields.many2one(
            'account.account',
            'Account',
            required=False,
            domain=[('parent_id', '!=', False)]),
        'type_report': fields.selection(
            [('sales_report', 'Sales Journal'),
             ('purchases_journal', 'Purchases Journal')],
            'Report Type',
            required=True),
    }

    _defaults = {
        'type_report': 'sales_report',
    }

    def print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}

        data = {
            'model': 'sales.purchases.journal.wizard',
            'ids': ids,
            'form': self.read(cr, uid, ids[0], context=context),
        }

        report_name = 'sales_journal_report'
        name = 'Sales Journal Report'

        if data['form']['type_report'] == 'purchases_journal':
            report_name = 'purchases_journal_report'
            name = 'Purchases Journal Report'

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


sales_purchases_journal_wizard()
