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
