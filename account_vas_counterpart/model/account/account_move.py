# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, models


class AccountMove(models.Model):
    _inherit = "account.move"

    @api.multi
    def set_counterpart(self):
        """
        This function is set counterpart
        for in case of create account move by hand.
        """
        acc_move_line_obj = self.env['account.move.line']
        for acc_move in self:
            acc_move_line_debit = acc_move_line_obj.search([
                ('move_id', '=', acc_move.id),
                ('debit', '!=', 0.0),
                ('counter_move_id', '=', False)
            ])

            acc_move_line_credit = acc_move_line_obj.search([
                ('move_id', '=', acc_move.id),
                ('credit', '!=', 0.0),
                ('counter_move_id', '=', False)
            ])

            if acc_move_line_debit and acc_move_line_credit:
                if len(acc_move_line_debit) == 1:
                    acc_move_line_credit.write(
                        {'counter_move_id': acc_move_line_debit.id})
                elif len(acc_move_line_credit) == 1:
                    acc_move_line_debit.write(
                        {'counter_move_id': acc_move_line_credit.id})
                else:
                    acc_move_line_grouped = {}
                    # group by name and amount debit or credit
                    for line in acc_move.line_ids:
                        if (line.name, line.credit == 0.0 and
                                line.debit or
                                line.credit) not in acc_move_line_grouped:

                            acc_move_line_grouped.update({
                                (line.name, line.credit == 0.0 and
                                 line.debit or line.credit):
                                [(line.debit == 0.0 and
                                  'credit'or 'debit', line.id)]
                            })
                        else:
                            acc_move_line_grouped[
                                (line.name, line.credit == 0.0 and
                                 line.debit or line.credit)].append(
                                (line.debit == 0.0 and 'credit' or
                                 'debit', line.id))
                    acc_dr = []
                    acc_cr = []
                    for values in acc_move_line_grouped.itervalues():
                        # set counterpart for couple line
                        if len(values) == 2:
                            acc_move_line_obj.browse(values[0][1]).write(
                                {'counter_move_id': values[1][1]})
                        else:
                            for value in values:
                                if value[0] == 'debit':
                                    acc_dr.append(value[1])
                                else:
                                    acc_cr.append(value[1])
                    if acc_dr and acc_cr and acc_dr > acc_cr:
                        acc_move_line_obj.browse(acc_dr).write(
                            {'counter_move_id': acc_cr[0]})
                    if acc_dr and acc_cr and acc_cr > acc_dr:
                        acc_move_line_obj.browse(acc_dr).write(
                            {'counter_move_id': acc_cr[0]})
        return True

    @api.multi
    def reset_counterpart(self):
        """
        we will reset counterpart_id for account_move
        when write, dupplicate new record.
        """
        # reset counter_move_id
        account_move_line_obj = self.env['account.move.line']
        # in case someone uses write function with ids is not a list
        account_move_line_ids = account_move_line_obj.search(
            [('move_id', 'in', self._ids)])
        # set all related counter_move_id to NULL
        if account_move_line_ids:
            account_move_line_ids.write({'counter_move_id': False})
            return True
        return False

    @api.model
    def create(self, vals):
        account_move = super(AccountMove, self).create(vals)
        account_move.reset_counterpart()
        account_move.set_counterpart()
        return account_move

    @api.multi
    def write(self, vals):
        res = super(AccountMove, self).write(vals)
        self.reset_counterpart()
        self.set_counterpart()
        return res
