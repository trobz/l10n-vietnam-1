# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import fields, models


class BalanceSheetConfig(models.Model):

    _name = 'balance.sheet.config'
    _rec_name = 'item'

    item = fields.Char(
        'Item',
        size=256
    )
    code = fields.Char(
        'Code',
        size=6,
        required=True
    )
    is_inverted_result = fields.Boolean(
        'Has Inverted Result?',
        help="Get the inverted Result"
    )
    is_parenthesis = fields.Boolean(
        'Has Parenthesis Result?',
        help="The result will be put in the Parenthesis"
    )
    config_line_ids = fields.One2many(
        'balance.sheet.config.line',
        'config_id',
        'Lines'
    )
    group_by_partner = fields.Boolean(
        'Group by Partner?',
        default=False
    )
