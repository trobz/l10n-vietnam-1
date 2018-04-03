# -*- coding: utf-8 -*-
{
    'name': 'Account VAS Counterpart',
    'version': '1.0',
    'category': '',
    'description': """
    In VAS (VietNam Accouting System), this module will have to set
    counterpart for related journal items when generating an journal entry.
    There two main function:
        - set_counterpart
        - reset_counterpart
    """,
    'author': 'Trobz',
    'website': 'http://www.trobz.com',
    'depends': [
        # OpenERP Native Modules
        'account',
        'account_voucher',
    ],
    'data': [
        # DATA


        # VIEWS
        'views/account/account_move_view.xml',

        # WIZARDS

        # REPORTS

        # MENUS

        # FUNCTIONS

    ],
    'test': [],
    'demo': [],
    'installable': True,
    'active': False,
    'application': True,
}
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
