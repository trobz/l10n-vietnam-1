# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
from datetime import datetime
from openerp.addons.report_base_vn.report import report_base_vn
from babel.numbers import format_number


class Parser(report_base_vn.Parser):

    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)
        self.report_name = 'htkk_purchases_report'
        self.total = 0.0
        self.total_tax = 0.0
        self.localcontext.update({
            'get_format_number': self.get_format_number,
            'get_wizard_data': self.get_wizard_data,
            'get_company': self.get_company,
            'get_result_base_on_target': self.get_result_base_on_target,
            'total': self.get_total,
            'total_tax': self.get_total_tax,
        })
#         self.get_wizard_data()

    def get_format_number(self, number):
        if isinstance(number, str):
            return number
        return format_number(number, locale='en_US')

    def get_total(self):
        return self.get_format_number(self.total)

    def get_total_tax(self):
        return self.get_format_number(self.total_tax)

    def get_company(self):
        res = self.pool.get('res.users').browse(self.cr, self.uid, self.uid)
        name = res.company_id.name
        tax_id = res.company_id.vat
        return {'name': name, 'tax_id': tax_id}

    def get_result_base_on_target(self):
        """
        this funtion is get all result of target

        """
        period_obj = self.pool.get('account.period')
        data = self.localcontext['data']['form']
        from_date = False
        to_date = False
        if data['from_period_id'] and data['from_period_id'][0] and \
                data['to_period_id'] and data['to_period_id'][0]:
            from_date = period_obj.read(
                self.cr, self.uid, data['from_period_id'][0], [
                    'date_start'])['date_start']
            to_date = period_obj.read(
                self.cr, self.uid, data['to_period_id'][0], [
                    'date_stop'])['date_stop']
        invoice_obj = self.pool.get('account.invoice')
        if not from_date or not to_date:
            invoice_ids = []
        else:
            invoice_ids = invoice_obj.search(
                self.cr, self.uid, [
                    ('type', 'in', ['in_invoice', 'out_refund']),
                    ('creation_date',
                     '>=', from_date),
                    ('creation_date',
                     '<=', to_date),
                    ('state', 'not in',
                        ['draft', 'cancel', 'proforma', 'proforma2'])])
        # save total of every tax_type
        total_1, total_2, total_3, total_4, total_5 = 0.0, 0.0, 0.0, 0.0, 0.0
        # save tax value of every tax_type
        total_tax_1, total_tax_2, total_tax_3, \
            total_tax_4, total_tax_5 = 0.0, 0.0, 0.0, 0.0, 0.0
        # list item of every tax_type in report
        list_item_1, list_item_2, list_item_3, \
            list_item_4, list_item_5 = [], [], [], [], []
        # serial of every tax_type
        no_1, no_2, no_3, no_4, no_5 = 1, 1, 1, 1, 1
        currency_obj = self.pool.get('res.currency')
        # customer invoice
        for invoice in invoice_obj.browse(self.cr, self.uid, invoice_ids):
            sub_total_1, sub_total_2, sub_total_3, \
                sub_total_4, sub_total_5 = 0.0, 0.0, 0.0, 0.0, 0.0
            sub_total_tax_1, sub_total_tax_2, sub_total_tax_3, \
                sub_total_tax_4, sub_total_tax_5 = 0.0, 0.0, 0.0, 0.0, 0.0
            is_vnd = invoice.currency_id.name == 'VND'
            item = {
                'template_number': invoice.template_number,
                'template_prefix': invoice.template_prefix,
                'prefix': invoice.prefix or '',
                'invoice_num': invoice.type == 'in_invoice' and
                invoice.supplier_invoice_number or
                invoice.customer_invoice_num,
                'invoice_date': invoice.date_invoice,
                'partner': invoice.partner_id.name,
                'partner_tax_id': invoice.partner_id.vat or '',
                'note': invoice.comment or '',
            }
            product_1, product_2, product_3, \
                product_4, product_5 = '', '', '', '', ''
            tax_type_1, tax_type_2, tax_type_3, \
                tax_type_4, tax_type_5 = '', '', '', '', ''
            price_subtotal_1, price_subtotal_2, price_subtotal_3, \
                price_subtotal_4, price_subtotal_5 = 0.0, 0.0, 0.0, 0.0, 0.0
            for line in invoice.invoice_line:
                sub_total = line.price_subtotal
                sub_total_tax = line.invoice_line_tax_id and \
                    line.price_subtotal * line.invoice_line_tax_id[
                        0].amount or 0.0
                # if invoice is not VND then compute them to VND
                if not is_vnd:
                    sub_total = currency_obj.compute(
                        self.cr, self.uid, invoice.currency_id.id,
                        invoice.company_id.currency_id.id,
                        line.price_subtotal,
                        context={'date': invoice.date_invoice})
                    sub_total_tax = currency_obj.compute(
                        self.cr, self.uid, invoice.currency_id.id,
                        invoice.company_id.currency_id.id,
                        sub_total_tax,
                        context={'date': invoice.date_invoice})
                tax_type = ','.join([str(int((tax.amount or 0) * 100))
                                     for tax in line.invoice_line_tax_id])
                if line.tax_type == 'tax_type_1':
                    sub_total_1 += sub_total
                    sub_total_tax_1 += sub_total_tax
                    if price_subtotal_1 < line.price_subtotal:
                        price_subtotal_1 = line.price_subtotal
                        product_1 = line.name
                        tax_type_1 = tax_type
                if line.tax_type == 'tax_type_2':
                    sub_total_2 += sub_total
                    sub_total_tax_2 += sub_total_tax
                    if price_subtotal_2 < line.price_subtotal:
                        price_subtotal_2 = line.price_subtotal
                        product_2 = line.name
                        tax_type_2 = tax_type
                if line.tax_type == 'tax_type_3':
                    sub_total_3 += sub_total
                    sub_total_tax_3 += sub_total_tax
                    if price_subtotal_3 < line.price_subtotal:
                        price_subtotal_3 = line.price_subtotal
                        product_3 = line.name
                        tax_type_3 = tax_type
                if line.tax_type == 'tax_type_4':
                    sub_total_4 += sub_total
                    sub_total_tax_4 += sub_total_tax
                    if price_subtotal_4 < line.price_subtotal:
                        price_subtotal_4 = line.price_subtotal
                        product_4 = line.name
                        tax_type_4 = tax_type
                if line.tax_type == 'tax_type_5':
                    sub_total_5 += sub_total
                    sub_total_tax_5 += sub_total_tax
                    if price_subtotal_5 < line.price_subtotal:
                        price_subtotal_5 = line.price_subtotal
                        product_5 = line.name
                        tax_type_5 = tax_type
            if product_1 != '':
                item_1 = item.copy()
                item_1.update({
                    'no': no_1,
                    'product': product_1,
                    'subtotal': sub_total_1,
                    'tax_type': tax_type_1,
                    'tax': sub_total_tax_1,
                })
                list_item_1.append(item_1)
                total_1 += sub_total_1
                total_tax_1 += sub_total_tax_1
                no_1 += 1
            if product_2 != '':
                item_2 = item.copy()
                item_2.update({
                    'no': no_2,
                    'product': product_2,
                    'subtotal': sub_total_2,
                    'tax_type': tax_type_2,
                    'tax': sub_total_tax_2,
                })
                list_item_2.append(item_2)
                total_2 += sub_total_2
                total_tax_2 += sub_total_tax_2
                no_2 += 1
            if product_3 != '':
                item_3 = item.copy()
                item_3.update({
                    'no': no_3,
                    'product': product_3,
                    'subtotal': sub_total_3,
                    'tax_type': tax_type_3,
                    'tax': sub_total_tax_3,
                })
                list_item_3.append(item_3)
                total_3 += sub_total_3
                total_tax_3 += sub_total_tax_3
                no_3 += 1
            if product_4 != '':
                item_4 = item.copy()
                item_4.update({
                    'no': no_4,
                    'product': product_4,
                    'subtotal': sub_total_4,
                    'tax_type': tax_type_4,
                    'tax': sub_total_tax_4,
                })
                list_item_4.append(item_4)
                total_4 += sub_total_4
                total_tax_4 += sub_total_tax_4
                no_4 += 1
            if product_5 != '':
                item_5 = item.copy()
                item_5.update({
                    'no': no_5,
                    'product': product_5,
                    'subtotal': sub_total_5,
                    'tax_type': tax_type_5,
                    'tax': sub_total_tax_5,
                })
                list_item_5.append(item_5)
                total_5 += sub_total_5
                total_tax_5 += sub_total_tax_5
                no_5 += 1

        # purchase reciept
        voucher_obj = self.pool.get('account.voucher')
        if not from_date or not to_date:
            voucher_ids = []
        else:
            voucher_ids = voucher_obj.search(
                self.cr, self.uid, [
                    ('type', '=', 'purchase'),
                    ('creation_date',
                     '>=', from_date),
                    ('creation_date',
                     '<=', to_date),
                    ('state', 'not in', ['draft', 'cancel', 'proforma'])])
        for voucher in voucher_obj.browse(self.cr, self.uid, voucher_ids):
            item = {}
            price_subtotal = 0.0
            product = ''
            for line in voucher.line_cr_ids:
                if line.amount > price_subtotal:
                    price_subtotal = line.amount
                    product = line.name
            item = {
                'template_number': voucher.template_number,
                'template_prefix': voucher.template_prefix,
                'prefix': voucher.prefix or '',
                'invoice_num': voucher.reference,
                'invoice_date': voucher.date,
                'partner': voucher.partner_id.name,
                'partner_tax_id': voucher.partner_id.vat or '',
                'product': product,
                'subtotal': voucher.amount - voucher.tax_amount,
                'tax_type': str(int((voucher.tax_id.amount or 0) * 100)),
                'tax': voucher.tax_amount,
                'note': voucher.narration or '',
            }
            if voucher.tax_type == 'tax_type_1':
                item.update({
                    'no': no_1,
                })
                list_item_1.append(item)
                total_1 += item['subtotal']
                total_tax_1 += item['tax']
                no_1 += 1
                continue
            if voucher.tax_type == 'tax_type_2':
                item.update({
                    'no': no_2,
                })
                list_item_2.append(item)
                total_2 += item['subtotal']
                total_tax_2 += item['tax']
                no_2 += 1
                continue
            if voucher.tax_type == 'tax_type_3':
                item.update({
                    'no': no_3,
                })
                list_item_3.append(item)
                total_3 += item['subtotal']
                total_tax_3 += item['tax']
                no_3 += 1
                continue
            if voucher.tax_type == 'tax_type_4':
                item.update({
                    'no': no_4,
                })
                list_item_4.append(item)
                total_4 += item['subtotal']
                total_tax_4 += item['tax']
                no_4 += 1
                continue
            if voucher.tax_type == 'tax_type_5':
                item.update({
                    'no': no_5,
                })
                list_item_5.append(item)
                total_5 += item['subtotal']
                total_tax_5 += item['tax']
                no_5 += 1
                continue
        self.total = total_1 + total_2 + total_3 + total_4 + total_5
        self.total_tax = total_tax_1 + total_tax_2 + \
            total_tax_3 + total_tax_4 + total_tax_5
        return [{('tax_type_1', total_1, total_tax_1): list_item_1},
                {('tax_type_2', total_2, total_tax_2): list_item_2},
                {('tax_type_3', total_3, total_tax_3): list_item_3},
                {('tax_type_4', total_4, total_tax_4): list_item_4},
                {('tax_type_5', total_5, total_tax_5): list_item_5}]

    def get_wizard_data(self):
        data = self.localcontext['data']['form']
        date_start_str = self.pool.get('account.period').read(
            self.cr, self.uid, data['from_period_id'][0], ['date_start'])
        date_start = datetime.strptime(
            date_start_str['date_start'], DEFAULT_SERVER_DATE_FORMAT)
        return {'month': date_start.month, 'year': date_start.year}
