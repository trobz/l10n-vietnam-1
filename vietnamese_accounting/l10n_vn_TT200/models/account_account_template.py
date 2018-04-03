# -*- coding: utf-8 -*-
from openerp import fields, models


class AccountAccountTemplate(models.Model):
    _inherit = "account.account.template"
    name = fields.Char(required=True, index=True, translate=True)

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
