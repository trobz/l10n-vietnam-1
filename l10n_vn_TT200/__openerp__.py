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
    "name": "Vietnam - Accounting",
    "version": "1.0",
    "author": "Trobz",
    'website': 'http://trobz.com',
    "category": "Localization/Account Charts",
    "description": """
Vietnamese Chart of Accounts
============================

This Chart of Accounts is based on Vietnamese Accounting Standard (VAS) under Circular No. 200/2014/TT-BTC.

The application also provides a simple wizard which is used to update accounts translation (into Vietnamese).

Installation:
-------------
* Note that this module conflicts with l10n_vn, just install only one of them.

Usage:
------
To use the wizard described above, you need to:

* Go to Accounting > Adviser > Chart of Accounts
* Select accounts that you want to update their translation (select on tree view)
* Click on "Action" button, which is on the right side of "Create" button
* Select "Update Chart of Accounts Translation", then select Company and Language on popped up wizard
* Click on "Update" button to start the translation process



""",
    "depends": ["account", "base_vat", "base_iban"],
    "data": [
        # DATA
        "data/account_chart.xml",
        "data/account_tax.xml",
        "data/account_chart_template.yml",

        # WIZARDS
        "wizards/update_account_translation_wizard.xml"
    ],
    "demo": [],
    "installable": True,
}
