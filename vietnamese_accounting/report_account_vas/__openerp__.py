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


{
    "name": "Report Accounting Vietnamese Localization",
    "version": "1.1",
    "author": "Trobz",
    "category": 'Generic Modules/Accounting',
    "description": """
Improve Accounting Management to map with Vietnamese Accounting
===============================================================
* VAS COMPLIANT PDF INVOICES
  - Reports generated in 3 copies
  - Reports archived in the system

* HTTK REPORT
  - HTTK reports for sales and purchases

* STOCK REPORT
  - Input note of material/ Tools & spare part/ Finish goods

* VIETNAMESE LANGUAGE
  - Vietnamese translation of the user interface of accounting modules
    """,
    'website': 'http://trobz.com',
    'init_xml': [],
    "depends": [
        'account',
        'stock',
        'report_base_vn',
        'post_function_one_time',
        'account_asset',
        'report_xls',
        'report_xlsx',
        'l10n_vn_TT200',
        'account_vas_counterpart',

    ],
    'data': [
        # data
        #         'data/properties_data.xml',
        #         'data/ir_config_account_data.xml',
        'data/cash_flow_direct_config_data.xml',

        # wizard
        'wizard/accounting/account_balance_sheet_view.xml',
        'wizard/accounting/account_payable_receivable_balance_view.xml',
        'wizard/accounting/account_report_profit_and_loss_view.xml',
        'wizard/accounting/account_report_general_journal_wizard_view.xml',
        'wizard/accounting/account_cash_book_view.xml',
        'wizard/accounting/account_asset_book_view.xml',
        'wizard/accounting/account_ledger_wizard_view.xml',
        'wizard/accounting/general_ledger_view.xml',
        'wizard/accounting/account_cash_flow_report_wizard_view.xml',
        'wizard/accounting/account_cash_flow_indirect_wizard.xml',
        'wizard/accounting/account_cash_flow_direct_wizard.xml',
        'wizard/accounting/account_stock_balance_wizard_view.xml',
        'wizard/accounting/account_sales_purchases_journal_wizard_view.xml',
        'wizard/accounting/account_receipt_payment_journal_wizard_view.xml',
        'wizard/accounting/account_stock_ledger_wizard_view.xml',
        'wizard/accounting/account_foreign_receivable_payable_ledger_view.xml',
        'wizard/accounting/account_report_trial_balance_view.xml',
        'wizard/accounting/print_htkk_purchases_wizard.xml',
        'wizard/accounting/print_htkk_sales_wizard.xml',
        'wizard/accounting/cash_bank_book_wizard.xml',
        'wizard/accounting/cash_book_wizard.xml',
        'wizard/accounting/asset_summary_report_wizard.xml',

        # report
        'report/financial_report/account_balance_sheet_report.xml',
        'report/financial_report/account_profit_and_loss_report.xml',
        'report/financial_report/account_cash_flow_direct_report.xml',
        'report/financial_report/account_profit_and_loss_xlsx_report.xml',
        'report/financial_report/financial_report.xml',

        'report/account_report.xml',

        'report/management_report/payable_receivable/general_payable_receivable_balance_report_view.xml',
        'report/management_report/payable_receivable/general_detail_payable_receivable_balance_report_view.xml',
        'report/management_report/payable_receivable/account_foreign_receivable_payable_ledger_report.xml',
        'report/management_report/account_receipt_journal.xml',


        'report/management_report/account_ledger_report.xml',
        'report/management_report/general_ledger_report.xml',
        'report/management_report/cash_book_report.xml',
        'report/management_report/cash_bank_book_report_xlsx.xml',
        'report/management_report/cash_book_report_xlsx.xml',
        'report/management_report/asset_book_report.xml',
        'report/management_report/account_stock_ledger_xlsx_report.xml',
        'report/management_report/asset_summary_xlsx_report.xml',

        'report/tax_report/htkk_purchases_report.xml',
        'report/tax_report/htkk_sales_report.xml',


        # view
        #         'view/account_invoice.xml',
        #         'view/account_voucher_view.xml',
        #         'view/account_view.xml',
        #         'view/account_asset_asset_view.xml',
        #         'view/account_move_line_view.xml',
        'view/account_move_view.xml',
        'view/account_account_view.xml',

        # ====== Configuration VAS Report =======#
        'view/config/balance_sheet_config_view.xml',
        'view/config/cash_flow_direct_config_view.xml',
        'view/config/profit_and_loss_config_view.xml',
        'view/config/indirect_cash_flow_config_view.xml',

        # menu
        'menu/accounting/accounting_menu.xml',
        'menu/config/vas_report_config_menu.xml',

        # edi
        #         'edi/invoice_action_data.xml',

        # POST INSTALL OBJECT
        'data/post_install_data.xml',
    ],
    'installable': True,
    'active': False,
    #     'post_init_hook': 'set_up_configurations_vas_report'
}
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
