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

from openerp import api, fields, models, _
from openerp.exceptions import UserError


class CashBookWizard(models.TransientModel):
    _name = 'cash.book.wizard'
    _inherit = "common.ledger"
    _description = 'Cash book'

    type_report = fields.Selection(
        [('cash_book', 'Cash book')],
        required=True, default='cash_book')
    journal_ids = fields.Many2many(
        'account.journal', string='Journals',
        required=False, default=False)
    target_move = fields.Selection(default='all')

    @api.multi
    def onchange_type(self, type_report):
        res = {'domain': {}}
        user = self.env.user
        account_code = "111%"
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
        report_name = 'cash_book_report_xlsx'
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
