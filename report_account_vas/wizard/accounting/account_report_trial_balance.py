# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import fields, osv
from openerp.tools.translate import _


class account_trial_balance_report(osv.osv_memory):
    _inherit = "account.common.account.report"
    _name = 'account.trial.balance.report'
    _description = 'Trial Balance Report'

    _columns = {
        'journal_ids': fields.many2many(
            'account.journal',
            'account_balance_report_journal_rel',
            'account_id',
            'journal_id',
            'Journals',
            required=True
        ),
    }

    _defaults = {
        'journal_ids': [],
        'display_account': 'not_zero',
    }

    def check_report(self, cr, uid, ids, context=None):
        if context is None:
            context = {}
        for data in self.browse(cr, uid, ids, context=context):
            date_from = False
            date_to = False
            if data.filter == 'filter_date':
                date_from = data.date_from or False
                date_to = data.date_to or False
                if (not date_from or not date_to) or (
                        date_from and date_to and date_from > date_to):
                    raise osv.except_osv(_('Warning!'), _(
                        'Date From must be less than Date To!'))
            elif data.filter == 'filter_period':
                if data.period_from.date_start > data.period_to.date_stop:
                    raise osv.except_osv(_('Warning!'), _(
                        'Date From must be less than Date To!'))
        res = super(account_trial_balance_report, self).check_report(
            cr, uid, ids, context=context)
        return res

    def _print_report(self, cr, uid, ids, data, context=None):
        data = self.pre_print_report(cr, uid, ids, data, context=context)
        return {'type': 'ir.actions.report.xml',
                'report_name': 'trial_balance_report',
                'name': 'Trial Balance Report',
                'datas': data}


account_trial_balance_report()
