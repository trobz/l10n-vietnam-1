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
