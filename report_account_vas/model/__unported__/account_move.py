# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import fields, osv


INVENTORY_VALUATION_PICKING_TYPE = [('purchase', 'Purchases'),
                                    ('sale', 'Sales')]


class account_move(osv.osv):
    _inherit = "account.move"

    _columns = {
        'description': fields.char(
            'Description',
            size=64
        ),
        'inventory_valuation_picking_type': fields.selection(
            INVENTORY_VALUATION_PICKING_TYPE,
            string='Inventory valuation picking type'
        ),
    }

    def fields_view_get(self, cr, uid, view_id=None, view_type=False,
                        context=None, toolbar=False, submenu=False):
        """
        override this function to remove <sheet> tag
        """
        res = super(account_move, self).fields_view_get(
            cr, uid, view_id=view_id, view_type=view_type,
            context=context, toolbar=toolbar, submenu=submenu
        )
        if '''<sheet string="Journal Entries">''' in res['arch'] \
                and "</sheet>" in res['arch']:
            res['arch'] = res['arch'] \
                .replace('''<sheet string="Journal Entries">''', '') \
                .replace('</sheet>', '')
        return res
