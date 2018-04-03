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


from openerp import models, api


class GeneralLedger(models.TransientModel):

    _name = 'general.ledger'
    _inherit = "common.ledger"
    _description = 'General Ledger'

    @api.multi
    def _print_report(self, data):
        if not data or not data.get('form', False):
            data = {
                'model': 'general.ledger',
                'ids': self.ids,
                'form': self.read(),
            }
        elif not data['form'].get('company_id', False) or \
                not data['form'].get('account_id', False):
            data['form']['company_id'] = self.read(['company_id'])[0]
            data['form']['account_id'] = self.read(['account_id'])[0]

        report_name = 'general_ledger_report_xls'
        name = 'General Ledger Report'

        return {
            'type': 'ir.actions.report.xml',
            'report_name': report_name,
            'datas': data,
            'name': name
        }


GeneralLedger()
