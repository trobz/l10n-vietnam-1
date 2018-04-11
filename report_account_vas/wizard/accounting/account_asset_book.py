# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.osv import fields, osv
from openerp.tools.translate import _
from openerp.exceptions import UserError


class account_cash_book(osv.osv_memory):

    _name = 'account.asset.book'
    _inherit = "account.common.report"
    _description = 'Asset Book Wizard Report'

    _columns = {
        'asset_category_id': fields.many2one('account.asset.category',
                                             'Asset Categories'),
    }

    def print_report(self, cr, uid, ids, data, context=None):
        # TODO: This report was implemented from Odoo version 7.0
        #     and it does not work in version 9.0
        raise UserError(_(
            "Sorry, this feature is not ready right now."
        ))
        if context is None:
            context = {}

        data = {
            'model': 'account.cash.book',
            'ids': ids,
            'form': self.read(cr, uid, ids[0], context=context),
        }
        report_name = 'asset.book.report'
        name = 'Asset Bank Report'
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


account_cash_book()
