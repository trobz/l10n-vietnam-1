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
#    along with this program. If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################


from openerp import fields, models


class balance_sheet_config(models.Model):
    _name = "balance.sheet.config.line"
    _order = 'code'

    code = fields.Char(
        'Account Code',
        required=True
    )
    config_id = fields.Many2one(
        'balance.sheet.config',
        'Config Line'
    )
    is_debit_balance = fields.Boolean(
        'Debit Balance',
        help='Get the debit balance of account'
    )
    is_credit_balance = fields.Boolean(
        'Credit Balance',
        help='Get the credit balance of account'
    )
    is_inverted = fields.Boolean(
        'Has Inverted?',
        help='When both debit and credit balance is checked, '
        'return negative balance if this field is checked.'
    )
    is_parenthesis = fields.Boolean(
        'Has parenthesis?',
        help='Return the negative of the account balance'
    )
    operator = fields.Selection(
        [
            ('plus', '+'),
            ('minus', '-'),
        ],
        'Operator',
        default='plus',
        required=True
    )
