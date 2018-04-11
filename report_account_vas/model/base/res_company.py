# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, fields, models


class ResCompany(models.Model):

    _inherit = 'res.company'

    @api.multi
    def get_company_info(self):
        # TODO: modify/inherit this function in order to get company info
        # Make sure 'name' and 'address' are listed in the dictionary
        self.ensure_one()
        info = {
            'name': self.name,
            'address': '',
            'vat': ''
        }

        return info
