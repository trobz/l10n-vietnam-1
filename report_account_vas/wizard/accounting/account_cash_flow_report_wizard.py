# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, models, fields
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
