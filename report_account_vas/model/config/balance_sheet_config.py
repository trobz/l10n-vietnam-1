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
