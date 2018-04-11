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

from openerp import api, models, fields


class AccountDetailedPayableReceivableBalance(models.TransientModel):

    _inherit = 'account.payable.receivable.balance'
    _name = 'account.detailed.payable.receivable.balance'
    _description = 'Print Detailed Payable And Receivable Report'

    partner_id = fields.Many2one('res.partner', 'Partner')

    journal_ids = fields.Many2many(
        comodel_name='account.journal',
        relation='account_detailed_payable_receivable_balance_journal_rel',
        column1='account_id', column2='journal_id',
        string='Journals', required=False)

    @api.onchange('account_type')
    def onchange_account_type(self):
        self.partner_id = False
        res = {}

        if self.account_type == 'receivable':
            res.update({
                'domain': {
                    'partner_id': [('customer', '=', True)]
                }
            })

        else:
            res.update({
                'domain': {
                    'partner_id': [('supplier', '=', True)]
                }
            })

        return res

    @api.multi
    def _print_report(self, data):
        report_name = 'general_detail_receivable_payable_balance'
        return self.env['report'].get_action(self, report_name, data=data)
