# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from sys import argv
import logging
from openerp.tools.translate import _
_logger = logging.getLogger(__name__)

to_19 = (u'không', u'một', u'hai', u'ba', u'bốn', u'năm', u'sáu',
         u'bảy', u'tám', u'chín', u'mười', u'mười một', u'mười hai',
         u'mười ba', u'mười bốn', u'mười lăm', u'mười sáu', u'mười bảy',
         u'mười tám', u'mười chín')
tens = (u'hai mươi', u'ba mươi', u'bốn mươi', u'năm mươi',
        u'sáu mươi', u'bảy mươi', u'tám mươi', u'chín mươi')
denom = ('',
         u'nghìn', u'triệu', u'tỷ', u'nghìn tỷ', u'trăm nghìn tỷ',
         'Quintillion', 'Sextillion', 'Septillion', 'Octillion', 'Nonillion',
         'Decillion', 'Undecillion', 'Duodecillion', 'Tredecillion',
         'Quattuordecillion', 'Sexdecillion', 'Septendecillion',
         'Octodecillion', 'Novemdecillion', 'Vigintillion')

# convert a value < 100 to English.


def _convert_nn(val):
    if val < 20:
        return to_19[val]
    for (dcap, dval) in ((k, 20 + (10 * v)) for (v, k) in enumerate(tens)):
        if dval + 10 > val:
            if val % 10:
                a = u'lăm'
#                if to_19[val % 10] == u'năm':
#                    a = u'lăm'
                if to_19[val % 10] == u'một':
                    a = u'mốt'
                else:
                    a = to_19[val % 10]
                return dcap + ' ' + a
            return dcap

# convert a value < 1000 to english, special cased because it is the level \
# that kicks
# off the < 100 special case.  The rest are more general.  This also allows \
# you to
# get strings in the form of 'forty-five hundred' if called directly.


def _convert_nnn(val):
    word = ''
    (mod, rem) = (val % 100, val // 100)
    if rem > 0:
        word = to_19[rem] + u' trăm'
        if mod > 0:
            word = word + ' '
    if mod > 0:
        word = word + _convert_nn(mod)
    return word


def vietnam_number(val):
    if val < 100:
        return _convert_nn(val)
    if val < 1000:
        return _convert_nnn(val)
    for (didx, dval) in ((v - 1, 1000 ** v) for v in range(len(denom))):
        if dval > val:
            mod = 1000 ** didx
            l = val // mod
            r = val - (l * mod)
            ret = _convert_nnn(l) + ' ' + denom[didx]
            if r > 0:
                ret = ret + ' ' + vietnam_number(r)
            return ret


def amount_to_text(number, currency):
    number = '%.2f' % number
    the_list = str(number).split('.')
    start_word = vietnam_number(int(the_list[0]))
    final_result = start_word
    if len(the_list) > 1 and int(the_list[1]) > 0:
        end_word = vietnam_number(int(the_list[1]))
        final_result = final_result + u' phẩy ' + end_word
    currency = currency == 'VND' and u'đồng' or currency
    return final_result + ' ' + currency

###############################################################################
# Generic functions
###############################################################################


_translate_funcs = {'vi': amount_to_text}


def amount_to_text_vi(nbr, lang='vi', currency='VND'):
    """ Converts an integer to its textual representation, using the language \
    set in the context if any.

        Example::

            1654: thousands six cent cinquante-quatre.
    """
    if lang not in _translate_funcs:
        _logger.warning(
            _("no translation function found for lang: '%s'"), lang)
        # TODO: (default should be en) same as above
        lang = 'en'
    return _translate_funcs[lang](abs(nbr), currency)


if __name__ == '__main__':
    # TODO: Duplicated code which is not compiling. Should it be deleted?
    lang = 'vi'
    if len(argv) < 2:
        for i in range(1, 200):
            print i, ">>", amount_to_text(i, lang)
        for i in range(200, 999999, 139):
            print i, ">>", amount_to_text(i, lang)
    else:
        print amount_to_text(int(argv[1]), lang)
