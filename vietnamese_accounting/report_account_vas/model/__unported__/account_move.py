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


INVENTORY_VALUATION_PICKING_TYPE = [('purchase', 'Purchases'),
                                    ('sale', 'Sales')]


class account_move(osv.osv):
    _inherit = "account.move"

    _columns = {
        'description': fields.char('Description', size=64),
        'inventory_valuation_picking_type': fields.selection(
            INVENTORY_VALUATION_PICKING_TYPE,
            string='Inventory valuation picking type'),
    }

    def fields_view_get(self, cr, uid, view_id=None, view_type=False,
                        context=None, toolbar=False, submenu=False):
        """
        override this function to remove <sheet> tag
        """
        res = super(account_move, self).fields_view_get(cr, uid,
                                                        view_id=view_id,
                                                        view_type=view_type,
                                                        context=context,
                                                        toolbar=toolbar,
                                                        submenu=submenu)
        if '''<sheet string="Journal Entries">''' in res['arch'] \
                and "</sheet>" in res['arch']:
            res['arch'] = res['arch'] \
                .replace('''<sheet string="Journal Entries">''', '') \
                .replace('</sheet>', '')
        return res
