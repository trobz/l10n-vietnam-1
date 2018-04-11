# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, models


class StockQuant(models.Model):
    _inherit = 'stock.quant'

    @api.model
    def _prepare_account_move_line(
            self, move, qty, cost, credit_account_id, debit_account_id):
        res = super(StockQuant, self)._prepare_account_move_line(
            move, qty, cost, credit_account_id, debit_account_id)

        res = self.modify_account_move_lines(res, move)
        return res

    @api.model
    def modify_account_move_lines(self, line_ids, move):
        for line in line_ids:
            vals = line[2]
            vals.update({'stock_move_id': move.id})

        return line_ids
