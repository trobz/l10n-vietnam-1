# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, fields, models


class AccountBalanceSheet(models.TransientModel):
    _name = "account.balance.sheet"
    _inherit = "account.common.account.report"
    _description = "Accounting Report"

    target_move = fields.Selection(default='posted')

    @api.multi
    def _print_report(self, data):
        report_name = 'balance_sheet_report'
        return self.env['report'].get_action(self, report_name, data=data)
