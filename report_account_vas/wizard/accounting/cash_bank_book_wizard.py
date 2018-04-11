# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, fields, models


class CashBankBookWizard(models.TransientModel):
    _name = 'cash.bank.book.wizard'
    _inherit = "common.ledger"
    _description = 'Cash at Bank book wizard'

    type_report = fields.Selection(
        [('cash_bank_book', 'Cash book at bank')],
        required=True, default='cash_bank_book')
    journal_ids = fields.Many2many(
        'account.journal', string='Journals',
        required=False, default=False)
    target_move = fields.Selection(default='all')

    @api.multi
    def onchange_type(self, type_report):
        res = {'domain': {}}
        user = self.env.user
        account_code = "112%"
        sql = """
            SELECT id
            FROM account_account
            WHERE internal_type = 'liquidity'
            AND code LIKE '%s'
            AND company_id = %d
        """ % (account_code, user.company_id.id)
        self.env.cr.execute(sql)
        chart_account_id = [i[0] for i in self.env.cr.fetchall()]
        res['domain'] = {'account_id': [('id', 'in', chart_account_id)]}
        return res

    @api.multi
    def _print_report(self, form_data):
        report_name = 'cash_bank_book_report_xlsx'
        if not self.journal_ids:
            journal_ids = self.env['account.journal'].search([])
            form = form_data.get('form')
            used_context = form.get('used_context')
            used_context.update({'journal_ids': journal_ids.ids})
            form.update({'journal_ids': journal_ids.ids,
                         'used_context': used_context})
            form_data.update({'form': form})
            self.journal_ids = journal_ids.ids
        return self.env['report'].get_action(self, report_name, data=form_data)
