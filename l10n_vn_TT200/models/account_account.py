# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).
from openerp import fields, models


class AccountAccount(models.Model):
    _inherit = "account.account"
    name = fields.Char(required=True, index=True, translate=True)
