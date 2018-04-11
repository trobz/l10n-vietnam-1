# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import fields, api, models, _
from datetime import date


class AccountProfitAndLossReportWizard(models.TransientModel):
    _inherit = "account.common.account.report"
    _name = 'account.profit.and.loss.report.wizard'
    _description = 'Profit and Loss Report'

    date_from = fields.Date(required=True, default=date.today())
    date_to = fields.Date(required=True, default=date.today())
    journal_ids = fields.Many2many(required=False)

    @api.multi
    def check_report(self):
        for record in self:
            if record.date_from > record.date_to:
                raise Warning(_("Start Date must be before End Date !"))

        report_name = 'account_profit_and_loss_report_xlsx'
        return self.env['report'].get_action(self, report_name)
