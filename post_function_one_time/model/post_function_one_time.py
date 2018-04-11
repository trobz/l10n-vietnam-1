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

from openerp import models, api
from openerp.tools.safe_eval import safe_eval
import logging

_logger = logging.getLogger(__name__)


class PostFunctionOneTime(models.AbstractModel):

    _name = "post.function.one.time"
    _description = "Post Function One Time"

    @api.model
    def run_post_object_one_time(self, object_name, list_functions=[]):
        """
        Generic function to run post object one time
        Input:
            + Object name: where you define the functions
            + List functions: to run
        Result:
            + Only functions which are not run before will be run
        """
        _logger.info(
            '==START running one time functions for post object: %s'
            % object_name
        )
        if isinstance(list_functions, (str, unicode)):
            list_functions = [list_functions]
        if not list_functions\
                or not isinstance(list_functions, (list)):
            _logger.warning('Invalid value of parameter list_functions.\
                            Exiting...')
            return False

        ir_conf_para_env = self.env['ir.config_parameter']
        post_object_env = self.env[object_name]
        ran_functions = \
            ir_conf_para_env.get_param(
                'List_post_object_one_time_functions', '[]')
        ran_functions = safe_eval(ran_functions)
        if not isinstance(ran_functions, (list)):
            ran_functions = []
        for function in list_functions:
            if (object_name + ':' + function) in ran_functions:
                continue
            getattr(post_object_env, function)()
            ran_functions.append(object_name + ':' + function)
        if ran_functions:
            ir_conf_para_env.set_param('List_post_object_one_time_functions',
                                       str(ran_functions))
        _logger.info('==END running one time functions for post object: %s'
                     % object_name)
        return True
