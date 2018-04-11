# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, fields, models, _
from openerp.exceptions import ValidationError


class CashFlowDirectConfig(models.Model):
    _name = "cash.flow.direct.config"
    _order = 'sequence'

    @api.model
    def _default_sequence(self):
        record = self.search([], order="sequence desc", limit=1)
        return record and record.sequence + 1 or 1

    name = fields.Char(
        string="Name",
        required=True
    )
    code = fields.Char(
        'Target Code',
    )
    type = fields.Selection(
        string="Type",
        selection=[
            ('title_only', 'Title Only'),
            ('detail_line', 'Detail Line'),
            ('total_line', 'Total Line'),
        ],
        default='title_only',
        required=True
    )
    data_period = fields.Selection(
        string="Data Period",
        selection=[
            ('within', 'Within Selected Period'),
            ('before', 'Before Selected Period'),
        ],
        default="within"
    )
    child_config_ids = fields.Many2many(
        string="Line(s) to Sum",
        comodel_name="cash.flow.direct.config",
        rel="parent_child_config_rel",
        column1="parent_id",
        column2="child_id"
    )
    sequence = fields.Integer(
        string="Sequence",
        default=lambda s: s._default_sequence(),
    )
    data_method = fields.Selection(
        string="Get Data Method",
        selection=[
            ('from_credit', 'Credit Only'),
            ('from_debit', 'Debit Only'),
            ('from_both', 'Both Credit and Debit'),
            ('from_either', 'Either Credit or Debit'),
        ],
        default='from_both',
    )
    compute_method = fields.Selection(
        string="Compute Amount Method",
        selection=[
            ('sum_debit', 'Amount = Total Debit'),
            ('sum_credit', 'Amount = Total Credit'),
            ('credit_debit', 'Amount = Total Credit - Total Debit'),
            ('debit_credit', 'Amount = Total Debit - Total Credit'),
        ],
        default='sum_debit',
    )
    credit_accounts = fields.Char(
        string="Credit Accounts",
    )
    debit_accounts = fields.Char(
        string="Debit Accounts",
    )
    advanced_credit_accounts = fields.Char(
        string="Credit Accounts",
    )
    advanced_debit_accounts = fields.Char(
        string="Debit Accounts",
    )
    description = fields.Text(
        string="Description",
        help="Value of this field will be printed in the third column"
        " of the report."
    )
    is_bold = fields.Boolean(
        string="Bold?",
        help='Check this if you want this line to be bold (font weight).'
    )
    has_parenthesis = fields.Boolean(
        string="Parenthesis?",
        help='Check this if you want to set result as negative value,'
        'it will be put inside parentheses by default. Example: (90,000),'
        ' which means the actual value is -90,000.'
    )

    @api.multi
    @api.constrains(
        'credit_accounts', 'debit_accounts',
        'advanced_credit_accounts', 'advanced_debit_accounts'
    )
    def _constraint_accounts(self):
        for config in self:
            if config.type == 'detail_line':
                error = False

                if not (
                    isinstance(config.credit_accounts, bool) or
                        config.credit_accounts.strip()
                ):
                    error = True
                if not (
                    isinstance(config.debit_accounts, bool) or
                        config.debit_accounts.strip()
                ):
                    error = True
                if not (
                    isinstance(config.advanced_credit_accounts, bool) or
                        config.advanced_credit_accounts.strip()
                ):
                    error = True
                if not (
                    isinstance(config.advanced_debit_accounts, bool) or
                    config.advanced_debit_accounts.strip()
                ):
                    error = True

                if error:
                    raise ValidationError(_(
                        'Make sure all listed accounts are separated by '
                        'commas!'
                    ))

    @api.multi
    def name_get(self):
        result = []
        for config in self:
            view_name = "%s%s" % (
                config.code and "[%s] " % config.code or '',
                config.name
            )
            result.append((config.id, view_name))

        return result
