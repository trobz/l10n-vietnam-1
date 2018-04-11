# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).
{
    "name": "Account VAS Counterpart",
    "version": "1.0",
    "category": "Accounting",
    "description": """
In VAS (Vietnam Accounting System), this module will set
counterpart for related journal items when generating a new journal entry.

There are two main functions:

* set_counterpart
* reset_counterpart
    """,
    "author": "Trobz, Odoo Community Association (OCA)",
    "website": "http://trobz.com",
    "depends": [
        # OpenERP Native Modules
        "account",
        "account_voucher",
    ],
    "data": [
        "views/account/account_move_view.xml",
    ],
    "installable": True,
    "active": False,
    "application": True,
}
