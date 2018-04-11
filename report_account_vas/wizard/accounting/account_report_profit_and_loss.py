# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import osv


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
