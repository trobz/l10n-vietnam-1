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


class account_asset_asset(osv.osv):
    _inherit = "account.asset.asset"

    _columns = {
        'delivery_number': fields.char(
            'Delivery Number',
            size=64,
            states={'draft': [('readonly', False)]}
        ),
        'code': fields.char(
            'Invoice Number',
            size=32,
            readonly=True
        ),
        'reference': fields.char(
            'Reference',
            size=32,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'delivery_date': fields.date(
            'Delivery Date',
            readonly=True, states={'draft': [('readonly', False)]}
        ),
        'basic_specification': fields.text(
            'Basic Specification',
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'manufacturer': fields.many2one(
            'res.country',
            'Manufacturer (country)',
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'manufacture_year': fields.char(
            'Manufacture Year',
            size=4,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'department': fields.char(
            'Department',
            size=64,
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'start_date_in_use': fields.date(
            'Start Year in use',
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'capacity': fields.char(
            'Capacity',
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'date_of_liquidation': fields.date(
            'Date of liquidation',
            readonly=True,
            states={'draft': [('readonly', False)]}
        ),
        'reason': fields.text(
            'Reason of liquidation',
            readonly=True,
            states={'draft': [('readonly', False)]}
        )

    }
