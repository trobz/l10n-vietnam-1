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
from openerp import tools


class PostObjectReportBaseVn(models.TransientModel):
    _name = 'post.object.report.base.vn'

    @api.model
    def start(self):
        self.update_value_report_url()
        return True

    @api.model
    def update_value_report_url(self):
        interface = tools.config.get('xmlrpc_interface', False)
        port = tools.config.get('xmlrpc_port', False)
        ir_config_para = self.env['ir.config_parameter']
        if interface and port:
            values = 'http://%s:%s' % (interface, str(port))
            ir_config_para.set_param('report.url', values)
        else:
            web_local = ir_config_para.get_param(key='web.base.url')
            ir_config_para.set_param('report.url', web_local)
        return True
