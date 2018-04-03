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


from datetime import datetime
from openerp.addons.report_base_vn.report import report_base_vn
from openerp.tools.safe_eval import safe_eval
from openerp.addons.report_xls.report_xls import report_xls
import xlwt
from openerp.addons.report_xls.utils import _render
from openerp.addons.report_xls.utils import rowcol_to_cell


# define mapping between
"""
TOTAL KEY là tập hợp các mã số chính trên báo cáo mà cần dữ liệu của nó
được lấy bằng cách cộng các chỉ tiêu con khác
ví dụ: 110: [111, 112]
diễn giải: mã số 110 là tổng của mẫ số 111 và 112
"""
KEY_TOTAL = {
    110: [111, 112],
    120: [121, 122, 123],
    130: [131, 132, 133, 134, 135, 136, 137, 139],
    140: [141, 149],
    150: [151, 152, 153, 154, 155],
    200: [210, 220, 230, 240, 250, 260],
    210: [211, 212, 213, 214, 215, 216, 219],
    220: [221, 224, 227],
    221: [222, 223],
    224: [225, 226],
    227: [228, 229],
    230: [231, 232],
    240: [241, 242],
    250: [251, 252, 253, 254, 255],
    260: [261, 262, 268],
    270: [100, 200],  # note
    300: [310, 330],  # note
    310: [311, 312, 313, 314, 315, 316, 317,
          318, 319, 320, 321, 322, 323, 324],
    330: [331, 332, 333, 334, 335, 336, 337,
          338, 339, 340, 341, 342, 343],
    400: [410, 430],
    410: [411, 412, 413, 414, 415, 416, 417,
          418, 419, 420, 421, 422],
    430: [431, 432, 433],
    440: [300, 400]
}


class balance_sheet_report_parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context=None):
        super(balance_sheet_report_parser, self).__init__(cr, uid, name,
                                                          context=context)
        self.report_name = 'balance_sheet_report'
        self.beginning_balance = 0.0
        self.ending_balance = 0.0
        self.beginning_balance_total = 0.0
        self.data = {}
        self.account_data = {}
        self.localcontext.update({
            'get_every_line': self.get_every_line,
            'get_date': self.get_date,
            'compute_data': self.compute_data,
            'company_info': {},
            'return_company_info': self.return_company_info,
            'return_beginning_balance': self.return_beginning_balance,
            'get_total': self.get_total,
            'return_beginning_balance_total':
            self.return_beginning_balance_total,
        })
        self.context = context
        self.get_company_info()

    def get_total(self, key):
        self.beginning_balance_total = 0.0
        ending_balance_total = 0.0
        for sub_key in KEY_TOTAL.get(key, []):
            self.get_every_line(sub_key)
            ending_balance_total += self.ending_balance
            self.beginning_balance_total += self.beginning_balance
        return ending_balance_total

    def set_context(self, objects, data, ids, report_type=None):
        new_ids = ids
        if (data['model'] == 'ir.ui.menu'):
            new_ids = 'chart_account_id' in data['form'] and \
                [data['form']['chart_account_id']] or []
            objects = self.pool.get('account.account').browse(self.cr,
                                                              self.uid,
                                                              new_ids)
        return super(balance_sheet_report_parser, self).\
            set_context(objects, data, new_ids, report_type=report_type)

    def return_beginning_balance_total(self):
        return self.beginning_balance_total

    def return_beginning_balance(self):
        return self.beginning_balance

    def return_company_info(self, key):
        """
        This function base on key to return value for this key.
        """
        return self.localcontext['company_info'].get(key, '')

    def get_company_info(self):
        """
        This function to compute company info include: name, address, currency
        result is save into company_info in localcontext
        """
        # TODO: Modify this function to get detailed address
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        address_list = [
            'STREET1',
            'STREET2',
            'CITY',
            'STATE NAME',
            'COUNTRY NAME',
        ]
        while '' in address_list:
            address_list.remove('')
        address = ', '.join(address_list)
        currency = res.company_id.currency_id.name
        self.localcontext['company_info'] = {'name': name,
                                             'address': address,
                                             'currency': currency}
        return True

    def get_date(self, data):
        date_to = data['form']['date_to']
        if not date_to:
            return u'ngày....tháng....năm....'
        date_to = datetime.strptime(date_to, '%Y-%m-%d')
        return u'ngày %d tháng %d năm %d' % (date_to.day, date_to.month,
                                             date_to.year)

    def get_every_line(self, key_line):
        """
        This funtion is compute data to return data for every row
        """
        account_obj = self.pool.get('account.account')
        balance_sheet_config_pool = self.pool.get('balance.sheet.config')
        if not key_line:
            return {}
        config_ids = balance_sheet_config_pool.\
            search(self.cr, self.uid, [('code', '=', key_line)])
        if config_ids:
            config_obj = balance_sheet_config_pool.browse(self.cr,
                                                          self.uid,
                                                          config_ids[0])
        else:
            return {}
        account_ids = [acc_id.id for acc_id in config_obj.account_ids]
        python_formular = config_obj.python_formular and \
            safe_eval(config_obj.python_formular) or False
        self.beginning_balance = 0.0
        self.ending_balance = 0.0
        if account_ids or python_formular:

            # caculate balance base on invoiving accounts
            for acc_id in account_ids:
                info_acc = self.account_data.get(acc_id, (0.0, 0.0, 0.0, 0.0))
                self.ending_balance += info_acc[2] - info_acc[3]
                self.beginning_balance += info_acc[0] - info_acc[1]

            # caculate balance base on python_formular
            if python_formular:
                for account_code, account_value in python_formular.items():
                    account_ids = account_obj.\
                        search(self.cr, self.uid, [('code', '=like',
                                                    '%s%%' % account_code)])
                    account_ending_balance = 0
                    account_beginning_balance = 0
                    for acc_id in account_ids:
                        # default we get debit amount
                        info_acc = self.account_data.get(acc_id,
                                                         (0.0, 0.0, 0.0, 0.0))
                        account_ending_balance += info_acc[2] - info_acc[3]
                        account_beginning_balance += info_acc[0] - info_acc[1]
                    if account_value.get('credit'):
                        # if get credit, multiply by -1
                        account_ending_balance *= -1
                        account_beginning_balance *= -1
                    # get operator
                    operator = account_value.get('operator', '+')
                    if operator == '-':
                        self.ending_balance -= account_ending_balance
                        self.beginning_balance -= account_beginning_balance
                    else:
                        self.ending_balance += account_ending_balance
                        self.beginning_balance += account_beginning_balance

            if config_obj.is_debit and not config_obj.is_credit:
                if self.ending_balance < 0:
                    self.ending_balance = 0
                if self.beginning_balance < 0:
                    self.beginning_balance = 0
            elif not config_obj.is_debit and config_obj.is_credit:
                self.ending_balance = self.ending_balance * -1
                if self.ending_balance < 0:
                    self.ending_balance = 0
                self.beginning_balance = self.beginning_balance * -1
                if self.beginning_balance < 0:
                    self.beginning_balance = 0

            if config_obj.is_inverted_result:
                self.beginning_balance *= -1
                self.ending_balance *= -1

            if config_obj.is_parenthesis:
                self.beginning_balance = abs(self.beginning_balance) * -1
                self.ending_balance = abs(self.ending_balance) * -1

        return self.ending_balance

    def compute_data(self, data):
        """
        compute data for all line
        result will be save to account_result in localcontext
        """
        form = data['form']
        ctx = self.context.copy()
        ctx['date_from'] = form['date_from']
        ctx['date_to'] = form['date_to']
        ctx['state'] = form['target_move']

        # we only need accounts which start with [0xx,1xx,2xx,3xx,4xx]
        account_pool = self.pool.get('account.account')
        account_ids = []
        account_types = [0, 1, 2, 3, 4]
        for acc_type in account_types:
            account_ids += account_pool.\
                search(self.cr, self.uid, [('code', '=like',
                                            '%s%%' % acc_type)])
        state = form['target_move'] == 'posted' and \
            " JOIN account_move m ON ml.move_id = m.id "\
            " WHERE m.state = 'posted' " or ' WHERE True '
        sql = '''
            SELECT ml.account_id
                ,COALESCE(SUM(CASE WHEN ml.date < '%s'
                    THEN debit END),0) AS debit_beginning
                ,COALESCE(SUM(CASE WHEN ml.date < '%s'
                    THEN credit END),0) AS credit_beginning
                ,COALESCE(SUM(CASE WHEN ml.date <= '%s'
                    THEN debit END),0) AS debit_ending
                ,COALESCE(SUM(CASE WHEN ml.date <= '%s'
                    THEN credit  END),0) AS credit_ending
            FROM account_move_line ml
             %s
            AND ml.account_id in %s
            GROUP BY ml.account_id;
        ''' % (ctx['date_from'], ctx['date_from'],
               ctx['date_to'], ctx['date_to'], state,
               tuple(account_ids + [-1, -1]))
        self.cr.execute(sql)
        for line in self.cr.fetchall():
            self.account_data.update({line[0]: (line[1], line[2],
                                                line[3], line[4])})
        return ''


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
