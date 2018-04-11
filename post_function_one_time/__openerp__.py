# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

{
    "name": "Post Function One Time",
    "version": "1.0",
    "author": "Trobz",
    "category": 'Tools',
    "description": """
This application supports the system to execute specific functions
only one time after the first upgrade. 

The executed functions will not be
recalled during next upgrades unless you remove them from the
"List_post_object_one_time_functions" system parameter.

There is one main function:

* run_post_object_one_time(object_name, list_functions=[])
    """,
    'website': 'http://trobz.com',
    'init_xml': [],
    "depends": [
        'web',
        'base'
    ],
    'data': [],
    'installable': True,
    'active': False,
}
