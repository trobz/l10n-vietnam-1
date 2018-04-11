# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


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
#             to_currency = obj.company_id and obj.company_id.currency_id or \
#					False
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
