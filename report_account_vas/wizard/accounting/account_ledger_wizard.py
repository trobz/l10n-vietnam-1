# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, fields, models, _
from openerp.exceptions import Warning
from datetime import date


class AccountLedgerWizard(models.TransientModel):
    _name = 'account.ledger.wizard'
    _inherit = 'common.ledger'
    _description = 'Account Ledger'

    date_from = fields.Date(required=True, default=date.today())
    date_to = fields.Date(required=True, default=date.today())
    journal_ids = fields.Many2many(required=False)

    @api.model
    def default_get(self, fields):
        res = super(AccountLedgerWizard, self).default_get(fields)

        res.update({'target_move': 'all'})
        return res

    @api.multi
    def print_report(self, data):
        for record in self:
            if record.date_from > record.date_to:
                raise Warning(_("Start Date must be before End Date !"))

        report_name = 'account_ledger_report_xlsx'
        return self.env['report'].get_action(self, report_name, data=data)
