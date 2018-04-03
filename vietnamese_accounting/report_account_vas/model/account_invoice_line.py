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


from openerp.osv import fields, osv


class account_invoice_line(osv.osv):
    _inherit = "account.invoice.line"
    _columns = {
        'tax_type': fields.selection([
            ('tax_type_1', '1'),
            ('tax_type_2', '2'),
            ('tax_type_3', '3'),
            ('tax_type_4', '4'),
            ('tax_type_5', '5'), ], 'Target', select=True),
    }

    def onchange_tax_id(self, cr, uid, ids, tax_ids, context=None):
        if not tax_ids[0][2] or context and context.get('type', '') in ['out_refund', 'in_invoice']:
            return {'value': {'tax_type': 'tax_type_1'}}
        tax = self.pool.get('account.tax').read(
            cr, uid, tax_ids[0][2][0], ['type', 'amount'])
        reference = {('percent', 0.0): 'tax_type_2',
                     ('percent', 0.05): 'tax_type_3',
                     ('percent', 0.1): 'tax_type_4',
                     }
        return {'value': {'tax_type': reference.get((tax['type'], tax['amount']), 'tax_type_1')}}

    def create(self, cr, uid, vals, context=None):

        inv_line_id = super(account_invoice_line, self).create(
            cr, uid, vals, context=context)
        if vals.get('invoice_line_tax_id', []):
            result = self.onchange_tax_id(
                cr, uid, [inv_line_id], vals['invoice_line_tax_id'], context=context)
            self.write(cr, uid, [inv_line_id], {
                       'tax_type': result['value']['tax_type']}, context=context)
        return inv_line_id
