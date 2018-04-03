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


from datetime import datetime
from dateutil.relativedelta import relativedelta
from openerp import fields, api, models, _


class AccountReportTrialBalanceWizard(models.TransientModel):
    _inherit = 'account.common.account.report'
    _name = 'account.report.trial.balance.wizard'
    _description = 'Trial Balance Report'

    journal_ids = fields.Many2many(required=False)
    display_account = fields.Selection(default='not_zero')
    chart_template_id = fields.Many2one('account.chart.template',
                                        string='Chart of Account',
                                        required=False,
                                        domain="[('visible','=', True)]")
    # TOREMOVE: For debug
    date_from = fields.Date(default=datetime.now() - relativedelta(days=60))
    date_to = fields.Date(default=datetime.now())

    @api.multi
    def check_report(self):
        for data in self:
            date_from = data.date_from or False
            date_to = data.date_to or False
            if (not date_from or not date_to) or \
                    (date_from and date_to and date_from > date_to):
                raise Warning(_('Date From must be less than Date To!'))

        return self.btn_generate_report()

    @api.multi
    def btn_generate_report(self):
        self.ensure_one()
        return self.env['report'].get_action(self, 'report_trial_balance_xlsx')
