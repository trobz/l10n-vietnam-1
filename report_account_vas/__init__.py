# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


import data
import report
import wizard
import model

"""
Temparory comment this code
Run post_object One-time instead of run post_hook
until config of report is OK
Run post_object One-time help us to easy run it again as we want by removing
function name from config parameter
"""
# from openerp import SUPERUSER_ID, api
#
#
# def set_up_configurations_vas_report(cr, registry):
#     env = api.Environment(cr, SUPERUSER_ID, {})
#     # Run config parameter for Balance Sheet report
#     env['post.object.report.account.vas']\
#         .set_balance_sheet_config_data_one_time()
#
#     # Run config parameter for Profit and Loss report
#     env['post.object.report.account.vas']\
#         .set_profit_and_loss_config_data_one_time()
#
#     # Run config parameter for Cash Flow indirect report
#     env['post.object.report.account.vas']\
#         .set_cash_flow_indirect_config_one_time()
