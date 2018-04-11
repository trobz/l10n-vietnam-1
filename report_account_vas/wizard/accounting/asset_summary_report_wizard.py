# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


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
