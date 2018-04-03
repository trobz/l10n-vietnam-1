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
from datetime import datetime
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT


class account_invoice(osv.osv):
    _inherit = "account.invoice"

    _columns = {
        'prefix': fields.char('Invoice Prefix', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'customer_invoice_num': fields.char('Customer Invoice Number', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'creation_date': fields.datetime('Creation Date', readonly=True),
        'template_number': fields.char('Invoice template number', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'template_prefix': fields.char('Invoice template prefix', size=64, required=False, readonly=True, states={'draft': [('readonly', False)]}),
    }

    _defaults = {
        'creation_date': lambda *a: datetime.now().strftime(DEFAULT_SERVER_DATETIME_FORMAT),

    }

    def delete_view_inherit(self, cr, uid):
        """
        delete view in mekongfurniture_module
        """
        view_obj = self.pool.get('ir.ui.view')
        view_ids = view_obj.search(
            cr, uid, [('name', '=', 'account.invoice.form.vas.inherit')])
        view_ids += view_obj.search(cr, uid,
                                    [('name', '=', 'account.invoice.supplier.form.vas.inherit')])
        if view_ids:
            view_obj.unlink(cr, uid, view_ids, context=None)
        return True
