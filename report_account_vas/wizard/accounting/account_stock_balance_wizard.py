# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.osv import fields, osv
from openerp.tools.translate import _


class account_stock_balance_wizard(osv.osv_memory):

    _name = 'account.stock.balance.wizard'
    _inherit = "common.ledger"
    _description = 'Print Stock Balance Wizard'

    _columns = {
        'account': fields.many2one(
            'account.account',
            'Account',
            required=True,
            domain=[
                ('parent_id', '!=', False),
                ('type', '=', 'other'),
                ('code', 'like', '15')
            ]
        ),
        'location_id': fields.many2one(
            'stock.location',
            'Location',
            required=True,
            domain=[('usage', '=', 'internal')]
        ),
    }

    def print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}

        data = {
            'model': 'account.move.line',
            'ids': ids,
            'form': self.read(cr, uid, ids[0], context=context),
        }

        report_name = 'stock_balance_report'
        name = 'Stock Balance Report'
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


account_stock_balance_wizard()
