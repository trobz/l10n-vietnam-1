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


from openerp.osv import fields, osv
from datetime import datetime, timedelta
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT


class account_profit_and_loss_report(osv.osv_memory):
    _inherit = "account.common.account.report"
    _name = 'account.profit.and.loss.report'
    _description = 'Profit and Loss Report'

    def _print_report(self, cr, uid, ids, data, context=None):
        return {'type': 'ir.actions.report.xml',
                'report_name': 'account_profit_and_loss_report_xls',
                'datas': data,
                'name': 'Profit and Lost'
                }


account_profit_and_loss_report()

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
