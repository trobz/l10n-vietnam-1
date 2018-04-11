# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import models, api, fields
from openerp.tools.translate import _


class CommonLedger(models.TransientModel):

    _name = 'common.ledger'
    _inherit = "account.common.report"
    _description = 'Common Ledger'

    account_id = fields.Many2one('account.account', 'Account', required=True)

    @api.constrains('date_from', 'date_to')
    def _check_dates(self):
        msg = _('Start Date must be before End Date !')
        for record in self:
            if record.date_from and record.date_to \
                    and record.date_from > record.date_to:
                raise Warning(msg)
        return True


CommonLedger()
