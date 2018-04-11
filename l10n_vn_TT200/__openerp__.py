# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

{
    "name": "Vietnam - Accounting",
    "version": "1.0",
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
To use the Update Translation Accounts wizard, you need to:

* Go to Accounting > Adviser > Chart of Accounts
* Select accounts that you want to update their translation (select on tree view)
* Click on "Action" button, which is on the right side of "Create" button
* Select "Update Chart of Accounts Translation", then select Company and Language on popped up wizard
* Click on "Update" button to start the translation process



""",
    "author": "Trobz, Odoo Community Association (OCA)",
    "website": "http://trobz.com",
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
