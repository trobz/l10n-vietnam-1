# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).


from openerp.report import report_sxw
from openerp import pooler
from .amount_to_text_vn import amount_to_text_vi as amount_to_test_v


class Parser(report_sxw.rml_parse):

    def __init__(self, cr, uid, name, context):
        super(Parser, self).__init__(cr, uid, name, context=context)
        # Initialize common variables
        self.partner_obj = pooler.get_pool(self.cr.dbname).get('res.partner')
        self.currency_obj = pooler.get_pool(self.cr.dbname).get('res.currency')
        self.amount_converted = False
        self.localcontext.update({
            'convert_amount': self.convert_amount,
            'addr_get': self.addr_get,
            'bank_get': self.bank_get,
            'amount_to_text_vi': self.amount_to_text_vi
        })

    def convert_amount(self, date, amount,
                       from_currency_id, to_currency_id=None):
        """
        Convert an amount between two currencies.
        """
        if not to_currency_id:
            # Get VND id
            to_currency_id = self.get_currency_id('VND')
        context = {'date': date}
        self.amount_converted = self.currency_obj.compute(
            self.cr, self.uid, from_currency_id, to_currency_id,
            amount, context=context)
        to_currency = \
            self.currency_obj.browse(
                self.cr, self.uid, to_currency_id, context=context)
        self.amount_converted = self.currency_obj.round(
            self.cr, self.uid, currency=to_currency,
            amount=self.amount_converted)
        return self.amount_converted

    def addr_get(self, partner_id, address_type='contact'):
        """
        Get the address information of a partner.
        """
        result = {
            'name': '',
            'street': '',
            'street2': '',
            'city': '',
            'country_id': '',
            'phone': '',
            'fax': '',
            'email': '',
            'address': ''
        }
        if not partner_id:
            return result
        partner_info = \
            self.partner_obj.browse(
                self.cr, self.uid, partner_id)
        if not partner_info:
            return result
        if not partner_info.street and not partner_info.street2 \
                and not partner_info.city and not partner_info.country_id  \
                and not partner_info.phone and not partner_info.fax \
                and not partner_info.email:
            # find address_type or first contact of partner
            contacts = partner_info.child_ids
            if contacts:
                partner_info = ''
                for contact in contacts:
                    if contact.type == address_type:
                        partner_info = contact
                        break
                if not partner_info:
                    partner_info = contacts[0]

        if partner_info:
            result['name'] = partner_info.name
            result['street'] = partner_info.street
            result['street2'] = partner_info.street2
            result['zip'] = partner_info.zip
            # partner_info.state_id and partner_info.state_id.name
            result['city'] = partner_info.city
            result['country_id'] = \
                partner_info.country_id and partner_info.country_id.name or ''
            result['phone'] = partner_info.phone
            result['fax'] = partner_info.fax
            result['email'] = partner_info.email
            result['address'] = self._addr_get_str(result)
        return result

    def _addr_get_str(self, address_info):
        """
        Return address information as a whole string.
        """
        if not address_info:
            return ''
        addr_str = ''
        if address_info['street']:
            addr_str += address_info['street'] + "\r\n"
        if address_info['street2']:
            addr_str += address_info['street2'] + "\r\n"
        # City and Country are on same line
        if address_info['zip']:
            addr_str += address_info['zip'] + ' '
        if address_info['city']:
            addr_str += address_info['city']
            if address_info['country_id']:
                addr_str += ' - '
            else:
                addr_str += "\r\n"
        if address_info['country_id']:
            addr_str += address_info['country_id'] + "\r\n"
        # Phone and Fax are on same line
        if address_info['phone']:
            addr_str += 'Phone' + ': ' + address_info['phone']
            if address_info['fax']:
                addr_str += '  '
            else:
                addr_str += "\r\n"
        if address_info['fax']:
            addr_str += 'Fax' + ': ' + address_info['fax'] + "\r\n"
        if address_info['email']:
            addr_str += 'Email' + ': ' + address_info['email']
        return addr_str

    def bank_get(self, bank_acc):
        """
        Get bank information of a bank account.
        @param bank_acc: recordset of res.partner.bank
        """
        result = {
            'name': '',
            'street': '',
            'street2': '',
            'city': '',
            'country': '',
            'bic': '',
            'bank_owner': '',
            'account_number': '',
            'owner_street': '',
            'owner_city': '',
            'owner_country': '',
            'bank_addr': '',
            'owner_addr': ''
        }
        result['name'] = bank_acc.bank_name or (
            bank_acc.bank_id and bank_acc.bank_id.name)
        result['street'] = bank_acc.bank_id.street or ''
        result['street2'] = bank_acc.bank_id.street2 or ''
        result['city'] = bank_acc.bank_id.city or ''
        result[
            'country'] = bank_acc.bank_id.country and \
            bank_acc.bank_id.country.name or ''
        result['bic'] = bank_acc.bank_id.bic or ''
        result[
            'bank_owner'] = bank_acc.partner_id and \
            bank_acc.partner_id.name or ''
        result['account_number'] = bank_acc.acc_number or ''
        result['owner_street'] = bank_acc.partner_id and \
            bank_acc.partner_id.street or ''
        result['owner_city'] = bank_acc.partner_id and \
            bank_acc.partner_id.city or ''
        result[
            'owner_country'] = bank_acc.partner_id and \
            bank_acc.partner_id.country_id and \
            bank_acc.partner_id.country_id.name or ''
        result['bank_addr'] = self._bank_get_addr(result)
        result['owner_addr'] = self._bank_get_addr({
            'street': result['owner_street'],
            'city': result['owner_city'],
            'country': result['owner_country'],
            'street2': '',
            'bic': ''
        })
        return result

    def _bank_get_addr(self, bank_info):
        """
        Get the address information of a partner's bank as a whole string.
        """
        if not bank_info:
            return ''
        addr_str = ''
        if bank_info.get('name'):
            addr_str += bank_info['name'] + '\r\n'
        if bank_info['street']:
            addr_str += bank_info['street'] + '\r\n'
        if bank_info['street2']:
            addr_str += bank_info['street2'] + '\r\n'
        if bank_info['city']:
            addr_str += bank_info['city']
            if bank_info['country']:
                addr_str += ' - '
            else:
                addr_str += '\r\n'
        if bank_info['country']:
            addr_str += bank_info['country'] + '\r\n'
        if bank_info.get('bic', False):
            addr_str += 'Swift' + ': ' + bank_info['bic']
        # remove enter line
        if addr_str[-2:] == '\r\n':
            addr_str = addr_str[:-2]
        return addr_str

    def amount_to_text_vi(self, nbr, lang='vi', currency='VND'):
        return amount_to_test_v(nbr, lang, currency)
