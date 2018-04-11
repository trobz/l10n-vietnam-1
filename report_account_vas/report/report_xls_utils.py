# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

import xlwt
from openerp.addons.report_xls.report_xls import report_xls
from datetime import datetime


class generic_report_xls_base(report_xls):

    def __init__(
            self, name, table, rml=False,
            parser=False, header=True, store=False
    ):
        super(generic_report_xls_base, self).__init__(
            name, table, rml, parser, header, store)

        # Set Default Cell Styles
        self.date_format = 'DD/MM/YYYY'
        self.decimal_format = '#,##0'

        # normal cell
        normal_cell_format = ['middle', 'center', 'wrap']
        self.style_date = self.get_cell_style(
            ['center', 'wrap', 'middle', 'borders_all'], self.date_format)
        self.style_date_right = self.get_cell_style(
            ['right', 'wrap', 'middle', 'borders_all'], self.date_format)
        self.style_decimal = self.get_cell_style(
            ['right', 'wrap', 'middle', 'borders_all'], self.decimal_format)
        self.style_decimal_bold = self.get_cell_style(
            ['right', 'wrap', 'middle', 'borders_all', 'bold'],
            self.decimal_format)
        self.normal_style = self.get_cell_style(normal_cell_format)
        self.normal_style_right = self.get_cell_style(
            ['middle', 'right', 'wrap', 'middle'])
        self.normal_style_left = self.get_cell_style(
            ['middle', 'left', 'wrap', 'middle'])
        self.normal_style_right_borderall = self.get_cell_style(
            ['middle', 'right', 'wrap', 'middle', 'borders_all'])
        self.normal_style_left_borderall = self.get_cell_style(
            ['middle', 'left', 'wrap', 'middle', 'borders_all'])
        self.normal_style_bold = self.get_cell_style(
            normal_cell_format + ['bold'])
        self.normal_style_italic = self.get_cell_style(
            normal_cell_format + ['italic'])
        self.normal_style_borderall = self.get_cell_style(
            normal_cell_format + ['borders_all'])
        self.normal_style_bold_borderall = self.get_cell_style(
            normal_cell_format + ['borders_all', 'bold'])

    def get_cell_style(self, styles, cell_style_format=None):
        """
        Get default style for the cell
        @param styles: list of style which want to use
        @rtype: rowstyle
        """
        res_style = ''
        self.xls_styles.update({'middle': 'align: vert center;'})
        for style in styles:
            res_style += self.xls_styles[style]
        return xlwt.easyxf(res_style, num_format_str=cell_style_format)

    def generate_xls_report(
            self, _p, _xs, data, objects, wb, report_name='Report Name'):
        """
        This function will generate excel report.
        We setup some configurations for the sheet by defaults
        @return: worksheet object
        """
        xls_footers = {
            'standard': (
                '&L&%(font_size)s&%(font_style)s' + datetime.now(
                ).strftime("%d/%m/%Y %H:%M:%S") +
                '&R&%(font_size)s&%(font_style)s&P / &N') % self.hf_params,
        }

        # set print sheet
        # cell_overwirte_ok => allow to merge cell in the sheet
        ws, _row_pos = self.generate_worksheet(
            _p, _xs, data, objects, wb, report_name, count=0)

        # set print header/footer
        ws.header_str = self.xls_headers['standard']
        ws.footer_str = xls_footers['standard']

        return ws

    def generate_worksheet(
            self, _p, _xs, data, objects, wb, report_name, count=0):
        """
        @summary: get new worksheet from workbook,
                  reset current row position in the new worksheet
        @return: new worksheet, new row_pos
        """
        report_name = count and (
            report_name[:31] + ' ' + str(count)) or report_name[:31]
        ws = wb.add_sheet(report_name, cell_overwrite_ok=True)
        ws.panes_frozen = True
        ws.remove_splits = True
        ws.portrait = 0  # Landscape
        ws.fit_num_pages = 1
        ws.fit_height_to_pages = 0
        ws.fit_width_to_pages = 1  # allow to print fit one page
        row_pos = 0

        return ws, row_pos
