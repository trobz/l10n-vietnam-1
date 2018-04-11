# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import fields, osv
from datetime import datetime
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT


class account_invoice(osv.osv):
    _inherit = "account.invoice"

    _columns = {
        'prefix': fields.char(
            'Invoice Prefix',
            size=64,
            required=False,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'customer_invoice_num': fields.char(
            'Customer Invoice Number',
            size=64,
            required=False,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'creation_date': fields.datetime(
            'Creation Date',
            readonly=True),
        'template_number': fields.char(
            'Invoice template number',
            size=64,
            required=False,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'template_prefix': fields.char(
            'Invoice template prefix',
            size=64,
            required=False,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
    }

    _defaults = {
        'creation_date': lambda *a: datetime.now().strftime(
            DEFAULT_SERVER_DATETIME_FORMAT),

    }

    def delete_view_inherit(self, cr, uid):
        """
        delete view in mekongfurniture_module
        """
        view_obj = self.pool.get('ir.ui.view')
        view_ids = view_obj.search(
            cr, uid, [('name', '=', 'account.invoice.form.vas.inherit')])
        view_ids += view_obj.search(
            cr, uid, [
                ('name', '=', 'account.invoice.supplier.form.vas.inherit')
            ])
        if view_ids:
            view_obj.unlink(cr, uid, view_ids, context=None)
        return True
