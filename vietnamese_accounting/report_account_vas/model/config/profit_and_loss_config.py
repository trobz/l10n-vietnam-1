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


from openerp.osv import fields, osv


class profit_and_loss_config(osv.osv):

    _name = "profit.and.loss.config"
    _rec_name = 'item'

    _columns = {
        'item': fields.char('Item', size=256),
        'code': fields.char('Code', size=6, required=True),
        'account_ids': fields.many2many('account.account', 'profit_loss_config_accounts_accounts_rel', 'profit_loss_config_id', 'account_id', string='Involving Accounts', help="Involving accounts to compute ?"),
        'counterpart_account_ids': fields.many2many('account.account', 'profit_loss_config_counterpart_accounts_accounts_rel', 'profit_loss_config_id', 'account_id', string='Counter-part Accounts', help="Involving counter-part accounts to compute ?"),
        'is_debit': fields.boolean('Debit Total', help='The amount will be get from debit ?'),
        'is_credit': fields.boolean('Credit Total', help='The amount will be get from credit ?'),
        'exception': fields.boolean('Exception', help='In special case, check this field to get Total Debit minus Total credit'),
        'note': fields.text('Note For Special Case')
    }
