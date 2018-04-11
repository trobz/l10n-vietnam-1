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
