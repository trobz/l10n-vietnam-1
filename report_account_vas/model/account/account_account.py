# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp import api, fields, models


class AccountAccount(models.Model):

    _inherit = 'account.account'

    parent_id = fields.Many2one(
        string='Parent Account',
        comodel_name='account.account'
    )
