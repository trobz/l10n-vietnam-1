# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import models, api, _


class VasAccountGeneralJournal(models.TransientModel):

    _inherit = "account.common.report"
    _name = "vas.account.general.journal"
    _description = "Accounting General Journal Report"

    @api.constrains('date_from', 'date_to')
    def _check_dates(self):
        msg = _('Start Date must be before End Date !')
        for record in self:
            if record.date_from and record.date_to \
                    and record.date_from > record.date_to:
                raise Warning(msg)
        return True

    @api.multi
    def _print_report(self, data):
        report_name = 'vas_account_general_journal_xlsx'
        return self.env['report'].get_action(self, report_name)


VasAccountGeneralJournal()
