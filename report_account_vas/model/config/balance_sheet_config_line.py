# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


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
