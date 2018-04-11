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
#    along with this program. If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

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
