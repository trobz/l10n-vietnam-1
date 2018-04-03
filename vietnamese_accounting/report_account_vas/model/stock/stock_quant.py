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
