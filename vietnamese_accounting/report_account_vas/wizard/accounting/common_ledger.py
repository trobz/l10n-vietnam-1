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
