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


from openerp import fields, models


class AccountMoveLine(models.Model):
    _inherit = "account.move.line"

    stock_move_id = fields.Many2one(
        comodel_name='stock.move',
        related='move_id.stock_move_id', string='Stock Move')

#     def _get_exchange_rate(self, cr, uid, ids, name, args, context=None):
#         if context is None:
#             context = {}
#         res_currency_pool = self.pool['res.currency']
#         res = {}.fromkeys(ids, 0.0)
#         for obj in self.browse(cr, uid, ids, context=context):
#             to_currency = obj.company_id and obj.company_id.currency_id or False
#             from_currency = obj.currency_id or False
#             if from_currency and to_currency:
#                 rate = res_currency_pool._get_conversion_rate(
#                     cr, uid,
#                     from_currency,
#                     to_currency,
#                     context=context)
#                 res[obj.id] = rate
#         return res

#     _columns = {
#         'exchange_rate': fields.function(
#             _get_exchange_rate,
#             type='float',
#             string='Exchange Rate',
#             help='The rate of the currency to the currency of rate 1'),
#     }
