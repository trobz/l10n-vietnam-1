# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


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
        if not tax_ids[0][2] or context and \
                context.get('type', '') in ['out_refund', 'in_invoice']:
            return {'value': {'tax_type': 'tax_type_1'}}
        tax = self.pool.get('account.tax').read(
            cr, uid, tax_ids[0][2][0], ['type', 'amount'])
        reference = {('percent', 0.0): 'tax_type_2',
                     ('percent', 0.05): 'tax_type_3',
                     ('percent', 0.1): 'tax_type_4',
                     }
        return {
            'value': {
                'tax_type': reference.get(
                    (tax['type'], tax['amount']), 'tax_type_1'
                )
            }
        }

    def create(self, cr, uid, vals, context=None):

        inv_line_id = super(account_invoice_line, self).create(
            cr, uid, vals, context=context)
        if vals.get('invoice_line_tax_id', []):
            result = self.onchange_tax_id(
                cr, uid, [inv_line_id],
                vals['invoice_line_tax_id'], context=context
            )
            self.write(
                cr, uid, [inv_line_id], {
                    'tax_type': result['value']['tax_type']
                }, context=context
            )
        return inv_line_id
