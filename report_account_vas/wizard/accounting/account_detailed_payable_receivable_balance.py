# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

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
