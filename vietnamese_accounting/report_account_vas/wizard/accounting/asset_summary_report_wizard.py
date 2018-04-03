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
#    along with this program.  If not see <http://www.gnu.org/licenses/>.
#
##############################################################################


from openerp import fields, models, api, _
from datetime import datetime


class AssetSummaryReportWizard(models.TransientModel):
    _name = 'asset.summary.report.wizard'

    company_id = fields.Many2one('res.company', string='Company',
                                 required=True,
                                 default=lambda self: self.env.user.company_id)
    from_date = fields.Date(
        string='From Date', default=lambda self: datetime.now(),
        required=True)
    to_date = fields.Date(
        string='To Date', default=lambda self: datetime.now(),
        required=True)

    @api.constrains('to_date')
    def check_to_date(self):
        if self.to_date < self.from_date:
            raise Warning(_('From Date cannot exceed To Date.'))

    @api.multi
    def btn_generate_report(self):
        self.ensure_one()
        return self.env['report'].get_action(
            self, 'asset_summary_xlsx_report',
            data={
                'ids': self.env.context.get('active_ids', [self.id]),
                'id': self.id,
                'form': self.read(
                    ['from_date', 'to_date', 'company_id'])}
        )
