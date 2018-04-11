# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


{
    "name": "Report Accounting Vietnamese Localization",
    "version": "1.1",
    "category": "Generic Modules/Accounting",
    "description": """
Improve Accounting Management to map with Vietnamese
===============================================================
This application provides different Vietnamese Standard Accounting Reports :

* Financial Reports
    - Trial Balance 
    - Profit & Los 
    - Cash Flow (Direct)
    - Balance Sheet
* Partner Reports
    - Receivable/Payable Ledger
    - Receivable/Payable Balance
* General Reports
    - General Journal
    - General Ledger
* Cash/Bank Reports
    - Cash Book
    - Cash at Bank Book
* Stock Reports
    - Stock Ledger

    """,
    "author": "Trobz, Odoo Community Association (OCA)",
    "website": "http://trobz.com",
    "init_xml": [],
    "depends": [
        "account",
        "stock",
        "report_base_vn",
        "post_function_one_time",
        "account_asset",
        "report_xls",
        "report_xlsx",
        "l10n_vn_TT200",
        "account_vas_counterpart",

    ],
    "data": [
        # data
        #         "data/properties_data.xml",
        "data/ir_config_account_data.xml",
        "data/cash_flow_direct_config_data.xml",

        # wizard
        "wizard/accounting/account_balance_sheet_view.xml",
        "wizard/accounting/account_payable_receivable_balance_view.xml",
        "wizard/accounting/account_report_profit_and_loss_view.xml",
        "wizard/accounting/account_report_general_journal_wizard_view.xml",
        "wizard/accounting/account_cash_book_view.xml",
        "wizard/accounting/account_asset_book_view.xml",
        "wizard/accounting/account_ledger_wizard_view.xml",
        "wizard/accounting/account_cash_flow_report_wizard_view.xml",
        "wizard/accounting/account_cash_flow_indirect_wizard.xml",
        "wizard/accounting/account_cash_flow_direct_wizard.xml",
        "wizard/accounting/account_stock_balance_wizard_view.xml",
        "wizard/accounting/account_sales_purchases_journal_wizard_view.xml",
        "wizard/accounting/account_receipt_payment_journal_wizard_view.xml",
        "wizard/accounting/account_stock_ledger_wizard_view.xml",
        "wizard/accounting/account_foreign_receivable_payable_ledger_view.xml",
        "wizard/accounting/account_report_trial_balance_view.xml",
        "wizard/accounting/print_htkk_purchases_wizard.xml",
        "wizard/accounting/print_htkk_sales_wizard.xml",
        "wizard/accounting/cash_bank_book_wizard.xml",
        "wizard/accounting/cash_book_wizard.xml",
        "wizard/accounting/asset_summary_report_wizard.xml",

        # report
        "report/financial_report/financial_report.xml",
        "report/account_report.xml",
        "report/management_report/payable_receivable/payable_receivable_report.xml",
        "report/management_report/management_report.xml",
        "report/tax_report/htkk_purchases_report.xml",
        "report/tax_report/htkk_sales_report.xml",

        # view
        #         "view/account_invoice.xml",
        #         "view/account_voucher_view.xml",
        #         "view/account_view.xml",
        #         "view/account_asset_asset_view.xml",
        #         "view/account_move_line_view.xml",
        "view/account_move_view.xml",
        "view/account_account_view.xml",

        # ====== Configuration VAS Report =======#
        "view/config/balance_sheet_config_view.xml",
        "view/config/cash_flow_direct_config_view.xml",
        "view/config/profit_and_loss_config_view.xml",
        "view/config/indirect_cash_flow_config_view.xml",

        # menu
        "menu/accounting/accounting_menu.xml",
        "menu/config/vas_report_config_menu.xml",

        # edi
        #         "edi/invoice_action_data.xml",

        # POST INSTALL OBJECT
        "data/post_install_data.xml",
    ],
    "installable": True,
    "active": False,
    #     "post_init_hook": "set_up_configurations_vas_report"
}
