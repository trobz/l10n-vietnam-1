# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.osv import fields, osv


class indirect_cash_flow_config(osv.osv):

    _name = "indirect.cash.flow.config"
    _rec_name = 'item'
    _order = 'create_date'

    def _get_child_ids(self, cr, uid, ids, field_name, arg, context=None):
        result = {}
        for record in self.browse(cr, uid, ids, context=context):
            sql = '''
            SELECT id
            FROM indirect_cash_flow_config
            WHERE parent_id = %s
            ''' % record.id
            cr.execute(sql)
            child_ids = [x[0] for x in cr.fetchall()]
            result[record.id] = child_ids
        return result

    _columns = {
        'item': fields.char('Item', size=256),
        'code': fields.char('Code', size=6, required=True),
        'parent_id': fields.many2one(
            'indirect.cash.flow.config', 'Parent ID'),
        'dr_account_ids': fields.many2many(
            'account.account',
            'indirect_cashflow_config_dr_accounts_rel',
            'config_id',
            'account_id',
            string='Accounts to get Debit Balance',
            help="Accounts to get Debit Balance ?"),
        'cr_account_ids': fields.many2many(
            'account.account',
            'indirect_cashflow_config_cr_accounts_rel',
            'config_id',
            'account_id', string='Accounts to get Credit Balance',
            help="Accounts to get Credit Balance ?"),
        'is_inverted_result': fields.boolean(
            'Has Inverted Result', help="Get the inverted Result"),
        'is_formula_active': fields.boolean(
            'Formula Active'),
        'child_ids': fields.function(
            _get_child_ids, type='many2many',
            relation="indirect.cash.flow.config",
            string="Child Config"),
        # section for formular active
        'account_ids': fields.many2many(
            'account.account',
            'indirect_cashflow_config_accounts_rel', 'config_id',
            'account_id', string='Accounts For Formula Active'),
        'is_positive_difference': fields.boolean(
            'Positive Difference'),
        'python_formular': fields.text(
            'Python Formular',
            help='Python Code Mapping for special case')
    }
