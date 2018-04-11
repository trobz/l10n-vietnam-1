# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, fields, models


class UpdateAcccounTranslationWizard(models.Model):
    _name = 'update.account.translation.wizard'

    company_id = fields.Many2one('res.company', 'Company', required=True)
    lang_id = fields.Many2one('res.lang', string='Language', required=True)

    @api.multi
    def button_update_account_name(self):
        for update_account_obj in self:
            company_id = update_account_obj.company_id.id or False
            lang = update_account_obj.lang_id and \
                update_account_obj.lang_id.code or 'en_US'
            translation_term = self.env['ir.translation']
            account_account = self.env['account.account']
            account_template = self.env['account.account.template']
            ids_account_template = account_template.search([])
            if ids_account_template:
                for account_template_line in ids_account_template:
                    ids_account = account_account.search(
                        [('company_id', '=', company_id),
                         ('code', 'in',
                                  [str(account_template_line.code),
                                   str(account_template_line.code).ljust(
                                      6, '0')])]) or None
                    ids_account_template_term = translation_term.search(
                        [('name', '=', 'account.account.template,name'),
                         ('res_id', '=', account_template_line.id),
                         ('lang', '=', lang)])
                    if ids_account and ids_account_template_term:
                        for id_account in ids_account:
                            ob_account_template_term = \
                                ids_account_template_term[0]
                            ids_account_term = translation_term.search(
                                [('name', '=', 'account.account,name'),
                                 ('res_id', '=', id_account.id)])
                            if ids_account_term:
                                ids_account_term.write(
                                    {'value': ob_account_template_term.value})
                            else:
                                translation_term.create({
                                    'name': 'account.account,name',
                                    'type': 'model',
                                    'lang': 'vi_VN',
                                    'res_id': id_account.id,
                                    'value': ob_account_template_term.value,
                                    'src': ob_account_template_term.src})
            view_id = self.env.ref(
                'l10n_vn_TT200.'
                'view_update_account_translation_wizard_form_successed')
            value = {
                'name': 'Updating Translation Successful',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'update.account.translation.wizard',
                'view_id': view_id.id,
                'type': 'ir.actions.act_window',
                'target': 'new',
            }

            return value
