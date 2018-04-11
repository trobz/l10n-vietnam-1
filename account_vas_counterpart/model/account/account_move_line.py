# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, fields, models


class account_move_line(models.Model):
    _inherit = "account.move.line"

    name = fields.Char('Name', size=255, required=True)
    counter_move_id = fields.Many2one(
        'account.move.line', 'Counterpart', required=False)
