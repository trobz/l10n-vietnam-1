# -*- coding: utf-8 -*-

from openerp import api, fields, models


class account_move_line(models.Model):
    _inherit = "account.move.line"

    name = fields.Char('Name', size=255, required=True)
    counter_move_id = fields.Many2one(
        'account.move.line', 'Counterpart', required=False)
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
