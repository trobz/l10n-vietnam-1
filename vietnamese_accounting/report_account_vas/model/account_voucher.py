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

from openerp.osv import osv, fields
from datetime import datetime
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT


class account_voucher(osv.osv):

    _inherit = "account.voucher"

    _columns = {
        'prefix': fields.char('Prefix', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'invoice_num': fields.char('Invoice Number', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'creation_date': fields.datetime('Creation Date', readonly=True),
        'date': fields.date('Invoiced Date', readonly=True, select=True, states={'draft': [('readonly', False)]}, help="Effective date for accounting entries"),
        'template_number': fields.char('Invoice template number', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'template_prefix': fields.char('Invoice template prefix', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'tax_type': fields.selection([
            ('tax_type_1', '1'),
            ('tax_type_2', '2'),
            ('tax_type_3', '3'),
            ('tax_type_4', '4'),
            ('tax_type_5', '5')], 'Target', select=True)
    }
    _defaults = {
        'creation_date': lambda *a: datetime.now().strftime(DEFAULT_SERVER_DATETIME_FORMAT),
    }

    def onchange_price(self, cr, uid, ids, line_ids, tax_id, partner_id=False, context=None):
        """
        update target_type base on amount of tax
        """

        res = super(account_voucher, self).onchange_price(
            cr, uid, ids, line_ids, tax_id, partner_id, context=context)
        if not tax_id:
            res.get('value', {}).update({'tax_type': 'tax_type_1'})
            return res
        tax = self.pool.get('account.tax').read(
            cr, uid, tax_id, ['type', 'amount'])
        reference = {('percent', 0.0): 'tax_type_2',
                     ('percent', 0.05): 'tax_type_3',
                     ('percent', 0.1): 'tax_type_4',
                     }
        res['value'].update({'tax_type': reference.get(
            (tax['type'], tax['amount']), 'tax_type_1')})
        return res
