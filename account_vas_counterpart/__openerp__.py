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
{
    'name': 'Account VAS Counterpart',
    'version': '1.0',
    'category': '',
    'description': """
In VAS (VietNam Accouting System), this module will have to set
counterpart for related journal items when generating an journal entry.

There two main function:

* set_counterpart
* reset_counterpart
    """,
    'author': 'Trobz',
    'website': 'http://www.trobz.com',
    'depends': [
        # OpenERP Native Modules
        'account',
        'account_voucher',
    ],
    'data': [
        'views/account/account_move_view.xml',
    ],
    'installable': True,
    'active': False,
    'application': True,
}
