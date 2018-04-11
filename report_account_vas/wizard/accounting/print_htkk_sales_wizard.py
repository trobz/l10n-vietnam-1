# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.osv import osv, fields


class print_htkk_sales_wizard(osv.osv_memory):
    _name = 'print.htkk.sales.wizard'
    _description = 'Wizard Print HTKK Sales'
    _columns = {
        'from_period_id': fields.many2one(
            'account.period',
            'From period',
            required=True
        ),
        'to_period_id': fields.many2one(
            'account.period',
            'To Period',
            required=True
        ),
    }

    def default_get(self, cr, uid, fields, context=None):
        '''
        set default value for period
        '''
        res = super(print_htkk_sales_wizard, self).default_get(
            cr, uid, fields, context=context)
        current_period = self.pool.get(
            'account.period').find(cr, uid, None, None)[0]
        if current_period - 1 < 0:
            res.update({'from_period_id': current_period,
                        'to_period_id': current_period})
        else:
            res.update({'from_period_id': current_period - 1,
                        'to_period_id': current_period - 1})
        return res

    def print_htkk_sales(self, cr, uid, ids, context=None):
        if context is None:
            context = {}
        data = {}
        data['id'] = context.get('active_ids', [])
        data['model'] = 'print.htkk.sales.wizard'
        data['form'] = self.read(cr, uid, ids)[0]
        return {
            'type': 'ir.actions.report.xml',
            'report_name': 'htkk_sales_report',
            'datas': data,
            'name': 'HTKK Sales'
        }
