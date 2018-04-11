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
#    along with this program. If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

{
    "name": "Post Function One Time",
    "version": "1.0",
    "author": "Trobz",
    "category": 'Administration',
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
