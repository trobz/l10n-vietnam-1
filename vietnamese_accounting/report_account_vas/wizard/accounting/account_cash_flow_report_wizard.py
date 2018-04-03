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


from openerp import api, models, fields
from openerp.tools.translate import _
from openerp.exceptions import UserError


class AccountCashFlowReportWizard(models.TransientModel):
    _inherit = 'account.common.account.report'
    _name = 'account.cash.flow.report.wizard'
    _description = 'Cash Flow Report'

    date_to = fields.Date(
        default=fields.Date.today()
    )

    @api.multi
    def check_report(self):
        for data in self:
            if data.date_from > data.date_to:
                raise UserError('Date From must be less than Date To!')

        res = super(AccountCashFlowReportWizard, self).check_report()

        return res

    @api.multi
    def _print_report(self, data):
        # TODO: We could use this for the "Cash Flow (Indirect)" also
        # Ex: Create new field -- type: selection('direct' or 'indirect')
        report_name = 'report_cash_flow_direct_xlsx'
        return self.env['report'].get_action(self, report_name, data=data)
