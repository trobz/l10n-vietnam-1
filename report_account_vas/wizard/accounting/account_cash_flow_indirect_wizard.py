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


from openerp import api, models
from openerp.tools.translate import _
from openerp.exceptions import ValidationError


class AccountCashFlowIndirectWizard(models.TransientModel):
    _inherit = 'account.common.account.report'
    _name = 'account.cash.flow.indirect.wizard'
    _description = 'Cash Flow Indirect'

    _columns = {

    }

    @api.multi
    def check_report(self):
        for record in self:
            date_from = False
            date_to = False
            date_from = record.date_from or False
            date_to = record.date_to or False
            if (not date_from or not date_to) or \
                    (date_from and date_to and date_from > date_to):
                raise ValidationError(
                    _('Date From must be less than Date To!'))
        res = super(AccountCashFlowIndirectWizard, self).check_report()
        return res

    def pre_print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}
        data['form'].update(self.read(cr, uid, ids, [], context=context)[0])
        return data

    def _print_report(self, cr, uid, ids, data, context=None):
        if context is None:
            context = {}
        data = self.pre_print_report(cr, uid, ids, data, context=context)
        data['id'] = context.get('active_ids', [])
        data['model'] = 'account.cash.flow.indirect.wizard'
        if context.get('report_type', '') == 'indirect':
            report_name = 'cash_flow_indirect_report'
            name = 'Cash Flow Indirect'
        else:
            report_name = 'account_cash_flow_direct_report_xls'
            name = 'Cash Flow direct'
        return {
            'type': 'ir.actions.report.xml',
            'report_name': report_name,
            'datas': data,
            'name': name,
        }
