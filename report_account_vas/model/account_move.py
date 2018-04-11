# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import fields, models


class AccountMove(models.Model):
    _inherit = "account.move"

    stock_move_id = fields.Many2one(
        comodel_name='stock.move', string='Stock Move')
