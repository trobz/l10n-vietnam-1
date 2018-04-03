# -*- coding: utf-8 -*-

{
    "name": "Vietnam - Accounting",
    "version": "1.0",
    "author": "Trobz",
    'website': 'http://trobz.com',
    "category": "Localization/Account Charts",
    "description": """
Vietnamese Chart of Accounts
============================

This module applies to companies based in Vietnamese Accounting Standard (VAS)
with Chart of account under Circular No. 200/2014/TT-BTC

""",
    "depends": ["account", "base_vat", "base_iban"],
    "data": [
             "account_chart.xml",
             "account_tax.xml",
             "account_chart_template.yml",
             "wizards/update_account_translation_wizard.xml"],
    "demo": [],
    "installable": True,
}
