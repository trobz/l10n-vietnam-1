# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, fields, models
from openerp.exceptions import UserError


class AccountPayableReceivableBalance(models.TransientModel):

    _inherit = 'account.common.partner.report'
    _name = 'account.payable.receivable.balance'
    _description = 'General Payable And Receivable Report'

    account_type = fields.Selection(
        selection=[
            ('receivable', 'Receivable Accounts'),
            ('payable', 'Payable Accounts')
        ],
        string="Account Type",
        required=True
    )

    account_id = fields.Many2one(
        'account.account',
        'Account',
        required=True
    )

    journal_ids = fields.Many2many(
        comodel_name='account.journal',
        relation='account_payable_receivable_balance_journal_rel',
        column1='account_id',
        column2='journal_id',
        string='Journals',
        default=False, required=False
    )
    target_move = fields.Selection(
        selection=[
            ('all', 'All Entries'),
            ('posted', 'All Posted Entries')
        ],
        string='Target Moves',
        required=True,
        default='all'
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Partner'
    )

    @api.multi
    def check_report(self):
        for data in self:
            date_from = data.date_from or False
            date_to = data.date_to or False
            if (not date_from or not date_to) or (
                    date_from and date_to and date_from > date_to):
                raise UserError('Date From must be less than Date To!')

        res = super(AccountPayableReceivableBalance, self).check_report()

        return res

    @api.multi
    def _print_report(self, data):
        report_name = 'general_receivable_payable_balance_xlsx'
        return self.env['report'].get_action(self, report_name, data=data)
