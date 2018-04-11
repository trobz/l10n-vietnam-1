# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).
from openerp import api, fields, models


class account_voucher(models.Model):
    _inherit = "account.voucher"

    @api.multi
    def set_counterpart_voucher(self):
        '''
        set couterpart for voucher
        '''
        acc_move_line_obj = self.env['account.move.line']
        for voucher in self:
            group_move_line = {}
            for move_line in voucher.move_ids:
                amount = move_line.debit or move_line.credit
                if (amount, move_line.name) not in group_move_line:
                    if move_line.debit != 0.0:
                        group_move_line.update(
                            {(amount, move_line.name): ([move_line.id], [])})
                    else:
                        group_move_line.update(
                            {(amount, move_line.name): ([], [move_line.id])})
                else:
                    if move_line.debit != 0.0:
                        group_move_line[(amount, move_line.name)][
                            0].append(move_line.id)
                    else:
                        group_move_line[(amount, move_line.name)][
                            1].append(move_line.id)
            debit_ids = []
            credit_ids = []
            for value in group_move_line.itervalues():
                if value[0] and value[1]:
                    acc_move_line_obj.browse(value[1]).write(
                        {'counter_move_id': value[0][0]})
                else:
                    debit_ids += value[0]
                    credit_ids += value[1]
            if debit_ids and credit_ids and len(debit_ids) > len(credit_ids):
                acc_move_line_obj.browse(debit_ids).write(
                    {'counter_move_id': credit_ids[0]})
            elif debit_ids and credit_ids \
                    and len(debit_ids) <= len(credit_ids):
                acc_move_line_obj.browse(credit_ids).write(
                    {'counter_move_id': debit_ids[0]})
        return True
