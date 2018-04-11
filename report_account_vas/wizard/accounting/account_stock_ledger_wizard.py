# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import pytz
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openerp import api, fields, models
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF


class AccountStockLedgerWizard(models.TransientModel):
    _name = 'account.stock.ledger.wizard'
    _inherit = "common.ledger"
    _description = 'Print Stock Ledger Wizard'

    account_id = fields.Many2one(
        comodel_name='account.account',
        string='Account', required=True)

    product_id = fields.Many2one(
        comodel_name='product.product',
        string='Product')

    product_tmpl_id = fields.Many2one(
        comodel_name='product.template',
        string='Product Template')

    location_id = fields.Many2one(
        comodel_name='stock.location',
        string='Location',
        required=True)

    date_from_dt = fields.Datetime(
        string='From Date in Datetime', compute='_compute_date_dt')

    date_to_dt = fields.Datetime(
        string='To Date in Datetime', compute='_compute_date_dt')

    @api.onchange('product_id')
    def onchange_product_id(self):
        if self.product_id:
            self.product_tmpl_id = self.product_id.product_tmpl_id

    @api.onchange('product_tmpl_id')
    def onchange_product_tmpl_id(self):
        if self.product_tmpl_id != self.product_id.product_tmpl_id:
            self.product_id = False
        if self.product_tmpl_id:
            return {
                'domain': {
                    'product_id': [
                        ('product_tmpl_id', '=', self.product_tmpl_id.id),
                    ]
                },
            }

        return {
            'domain': {'product_id': [], },
        }

    @api.depends('date_from', 'date_to')
    def _compute_date_dt(self):
        for record in self:
            date_from_dt_tz = datetime.strptime(record.date_from, DF) + \
                relativedelta(hour=0, minute=0, second=0)

            date_to_dt_tz = datetime.strptime(record.date_to, DF) + \
                relativedelta(hour=23, minute=59, second=59)

            self.date_from_dt = date_from_dt_tz
            self.date_to_dt = date_to_dt_tz

    @api.model
    def convert_datetime_to_utc(self, datetime_tz):
        current_tz = self.env.context.get('tz', False)
        if current_tz:
            time_zone = pytz.timezone(current_tz)
            datetime_tz = datetime_tz.replace(tzinfo=time_zone).\
                astimezone(pytz.utc)

        return datetime_tz

    @api.multi
    def print_report(self, data):
        report_name = 'stock_ledger_report'
        return self.env['report'].get_action(self, report_name, data=data)
