# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF

from datetime import datetime


class AssetSummaryReportXlsx(ReportXlsx):

    _COLUMNS = {'A': {'name': u'STT', 'width': 5},
                'B': {'name': u'Mã TS', 'width': 15},
                'C': {'name': u'Tên Tài Sản', 'width': 20},
                'D': {'name': u'Số Lượng', 'width': 10},
                'E': {'name': u'Ngày Bắt Đầu', 'width': 10},
                'F': {'name': u'Ngày Kết Thúc', 'width': 10},
                'G': {'name': u'Thời Gian Khấu Hao', 'width': 10},
                'H': {'name': u'Thời Gian Đã Khấu Hao', 'width': 20},
                'I': {'name': u'Thời Gian Còn Lại', 'width': 10},
                'J': {'name': u'Nguyên Giá', 'width': 10},
                'K': {'name': u'Giá Trị Đã Khấu Hao', 'width': 10},
                'L': {'name': u'Giá Trị Phát Sinh Tăng', 'width': 10},
                'M': {'name': u'Giá Trị Phát Sinh Giảm', 'width': 10},
                'N': {'name': u'Giá Trị Còn Lại Phải Phân Bổ', 'width': 20},
                'O': {'name': u'Tài Khoản Khấu Hao', 'width': 10},
                'P': {'name': u'Tài Khoản Phân Bổ', 'width': 10},
                'Q': {'name': u'Phòng Ban', 'width': 20},
                }

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]
        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self._set_default_format()

        # generate header
        self.generate_company_info()

        # generate main content
        self.generate_report_title()
        self.generate_content_header()
        self.generate_main_content()

    def _define_formats(self, workbook):
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        format_config = {
            'font_name': 'Times New Roman',
            'font_size': 12,
            'valign': 'vcenter',
            'text_wrap': True,
        }
        self.format_default = workbook.add_format(format_config)

        format_bold = format_config.copy()
        format_bold.update({
            'bold': True,
        })
        self.format_bold = workbook.add_format(format_bold)

        format_bold_center = format_bold.copy()
        format_bold_center.update({
            'align': 'center',
        })
        self.format_bold_center = workbook.add_format(format_bold_center)

        format_right = format_config.copy()
        format_right.update({
            'align': 'right',
        })
        self.format_right = workbook.add_format(format_right)

        format_center = format_config.copy()
        format_center.update({
            'align': 'center',
        })
        self.format_center = workbook.add_format(format_center)

        # ---------------------------------------------------------------------
        # Company Info
        # ---------------------------------------------------------------------
        format_company_info = format_config.copy()
        format_company_info.update({
            'bold': True,
        })

        self.format_company_info = workbook.add_format(
            format_company_info)

        # ---------------------------------------------------------------------
        # Report Template
        # ---------------------------------------------------------------------
        format_template_title = format_config.copy()
        format_template_title.update({
            'bold': True,
            'align': 'center',
        })
        self.format_template_title = workbook.add_format(format_template_title)

        format_template_desc = format_config.copy()
        format_template_desc.update({
            'italic': True,
            'align': 'center',
        })
        self.format_template_desc = workbook.add_format(
            format_template_desc)

        # ---------------------------------------------------------------------
        # Report Title
        # ---------------------------------------------------------------------
        format_report_title = format_config.copy()
        format_report_title.update({
            'bold': True,
            'align': 'center',
            'font_size': 14,
        })
        self.format_report_title = workbook.add_format(format_report_title)

        format_report_date = format_config.copy()
        format_report_date.update({
            'italic': True,
            'align': 'center',
        })
        self.format_report_date = workbook.add_format(format_report_date)

        format_uom = format_config.copy()
        format_uom.update({
            'italic': True,
            'align': 'right',
        })
        self.format_uom = workbook.add_format(format_uom)
        # ---------------------------------------------------------------------
        # Table format
        # ---------------------------------------------------------------------
        format_table = format_config.copy()
        format_table.update({
            'border': True,
            'num_format': '#,##0',
        })
        self.format_table = workbook.add_format(format_table)

        format_table_date = format_config.copy()
        format_table_date.update({
            'border': True,
            'align': 'right',
            'num_format': 'dd/mm/yyyy'
        })
        self.format_table_date = workbook.add_format(format_table_date)

        format_table_center = format_table.copy()
        format_table_center.update({
            'align': 'center',
        })
        self.format_table_center = workbook.add_format(format_table_center)

        format_table_right = format_table.copy()
        format_table_right.update({
            'align': 'right',
        })
        self.format_table_right = workbook.add_format(format_table_right)

        format_table_bold = format_table.copy()
        format_table_bold.update({
            'bold': True,
        })
        self.format_table_bold = workbook.add_format(format_table_bold)

        format_table_bold_center = format_table_bold.copy()
        format_table_bold_center.update({
            'align': 'center',
        })
        self.format_table_bold_center = workbook.add_format(
            format_table_bold_center)

        format_table_bold_right = format_table_bold.copy()
        format_table_bold_right.update({
            'align': 'right',
        })
        self.format_table_bold_right = workbook.add_format(
            format_table_bold_right)

    def _set_default_format(self):
        self.sheet.hide_gridlines(2)
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(4, 20)
        self.sheet.set_row(7, 35)

    def generate_company_info(self):
        info = self.object.env.user.company_id.get_company_info()
        info = u'{0}\n{1}'.format(
            info['name'],
            u'Địa chỉ: %s' % info['address'],
        )
        self.sheet.merge_range(
            'A1:Q3', info,
            self.format_company_info)

    def generate_report_title(self):
        self.sheet.merge_range(
            'A{0}:Q{0}'.format(5),
            u'BẢNG THEO DÕI TĂNG GIẢM TSCĐ',
            self.format_report_title)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A{0}:Q{0}'.format(6),
            u'Từ ngày %s đến ngày %s' % (date_info['from_date'],
                                         date_info['to_date']),
            self.format_center)

    def get_date_info(self):
        from_date = self.convert_date_format(self.object.from_date)
        to_date = self.convert_date_format(self.object.to_date)
        return {
            'from_date': from_date,
            'to_date': to_date
        }

    def convert_date_format(self, date_str):
        return datetime.strftime(
            datetime.strptime(date_str, DF), '%d/%m/%Y')

    def generate_content_header(self):
        s_row = self.row_position = 8

        for key, value in self._COLUMNS.iteritems():
            self.sheet.set_column('{0}:{0}'.format(key), value['width'])
            self.sheet.write('{0}{1}'.format(key, s_row), value['name'],
                             self.format_table_bold_center)

    # Data region
    def get_lines(self):
        params = self.get_params_for_query()

        sql = """
WITH asset AS (
    SELECT _ass.id, _ass.code, _ass.name,
        to_char(_ass.date, 'YYYY-MM-DD') AS start_date,
        to_char(_ass.end_date, 'YYYY-MM-DD') AS end_date,
        (CASE WHEN method_time = 'number' THEN _ass.method_number
              WHEN method_time = 'end' THEN (
              (DATE_PART('year', method_end) - DATE_PART('year', date)) * 12 +
              (DATE_PART('month', method_end) - DATE_PART('month', date)))
         END) AS method_number,
        _ass.category_id, _ass.value, _ass.qty_onhand,
        COALESCE(_assl.depreciated_value, 0) AS depreciated_value
    FROM account_asset_asset _ass
        LEFT JOIN (SELECT asset_id, SUM(amount) AS depreciated_value
              FROM account_asset_depreciation_line
              WHERE depreciation_date >= '%(from_date)s'
                AND depreciation_date < '%(to_date)s'
              GROUP BY asset_id
              ) _assl ON _ass.id = _assl.asset_id
    )
    , depreciated_info AS (
        SELECT asset.id,
            asset.code,
            asset.name,
            asset.start_date,
            asset.end_date,
            asset.method_number,
            asset.category_id,
            asset.value,
            asset.depreciated_value,
            asset.qty_onhand,

            SUM(_aml_db.debit) AS sum_debit,
            COALESCE(COUNT(DISTINCT _aml_db.move_id), 0) AS depreciated_num,
            (asset.method_number - COALESCE(COUNT(
                distinct _aml_db.move_id), 0)) AS depreciated_remain
        FROM asset
            JOIN account_asset_depreciation_line _assl
                ON asset.id = _assl.asset_id
            JOIN account_asset_depreciation_line_account_move_rel _assmv
                ON _assl.id = _assmv.account_asset_depreciation_line_id
            JOIN (SELECT * FROM account_move_line
                  WHERE debit > 0
                        AND date_maturity < '%(from_date)s'
                  )
                _aml_db ON _aml_db.move_id = _assmv.account_move_id
            JOIN account_move _amv ON _aml_db.move_id = _amv.id
        WHERE _amv.state = 'posted'
            AND _amv.date < '%(from_date)s'
        GROUP BY asset.id,
            asset.method_number,
            asset.code,
            asset."name",
            asset.start_date,
            asset.end_date,
            asset.method_number,
            asset.category_id,
            asset."value",
            asset.depreciated_value,
            asset.qty_onhand
        ORDER BY asset.category_id
    )
    , depreciating_info AS (
        SELECT asset.id,
            string_agg(DISTINCT(_aml_cr.account_id::text), ', ') AS dep_acc,
            string_agg(DISTINCT(_aml_db.account_id::text), ', ') AS app_acc
        FROM asset
            JOIN account_asset_depreciation_line _assl
                ON asset.id = _assl.asset_id
            JOIN account_asset_depreciation_line_account_move_rel _assmv
                ON _assl.id = _assmv.account_asset_depreciation_line_id
            JOIN (SELECT * FROM account_move_line
                  WHERE debit > 0
                        AND date_maturity >= '%(from_date)s'
                        AND date_maturity < '%(to_date)s'
                  )
                _aml_db ON _aml_db.move_id = _assmv.account_move_id
            JOIN account_move _amv ON _aml_db.move_id = _amv.id
            JOIN (SELECT * FROM account_move_line
                  WHERE credit > 0
                        AND date_maturity >= '%(from_date)s' AND
                        date_maturity < '%(to_date)s'
                  )
                _aml_cr ON _aml_cr.move_id = _aml_db.move_id
        WHERE _amv.state = 'posted'
            AND _amv.date >= '%(from_date)s' AND _amv.date < '%(to_date)s'
        GROUP BY asset.id
    )
    , summary AS (
        SELECT depreciated_info.*, dep_acc, app_acc
        FROM
            depreciated_info
            LEFT JOIN depreciating_info
                ON depreciated_info.id = depreciating_info.id
    )
    SELECT row_to_json (summary) FROM summary
        """
        sql = sql % params
        self.env.cr.execute(sql)
        res = self.env.cr.fetchall()

        # Get asset group
        asset_gps = {}
        ids = [data[0].get('category_id') for data in res]
        if ids:
            ids = 'IN {}'.format(tuple(list(set(ids)) + [-1, -1]))
            sql = """
                WITH temp AS (
                        SELECT id, name
                        FROM account_asset_category
                        WHERE id %(asset_gps)s
                    )
                    SELECT row_to_json(temp) FROM temp
            """ % {'asset_gps': ids}
            self.env.cr.execute(sql)

            for line in self.env.cr.fetchall():
                line = line[0]
                asset_gps[line.get('id')] = line.get('name')
        return res, asset_gps

    def get_params_for_query(self):
        return {
            'from_date': self.object.from_date,
            'to_date': self.object.to_date
        }

    def generate_main_content(self):
        s_row = self.row_position = 9

        # Data
        stt = 1
        lines_data, asset_gps = self.get_lines()
        bk_gp = 0
        for line in lines_data:
            line = line[0]
            category_id = line.get('category_id', 0)
            if bk_gp != category_id:
                bk_gp = category_id
                for col in self._COLUMNS.keys():
                    if col == 'C':
                        self.sheet.write(
                            'C{}'.format(s_row),
                            asset_gps.get(category_id, u'Unknown'),
                            self.format_table_bold)
                    else:
                        self.sheet.write('{0}{1}'.format(col, s_row), '',
                                         self.format_table)
                stt = 1
                s_row += 1

            self.sheet.write('A{}'.format(s_row), stt, self.format_table)
            self.sheet.write('B{}'.format(s_row), line.get('code'),
                             self.format_table)
            self.sheet.write('C{}'.format(s_row), line.get('name'),
                             self.format_table)
            self.sheet.write('D{}'.format(s_row), line.get('qty_onhand'),
                             self.format_table)
            self.sheet.write('E{}'.format(s_row), line.get('start_date'),
                             self.format_table_date)
            self.sheet.write('F{}'.format(s_row), line.get('end_date'),
                             self.format_table_date)
            self.sheet.write('G{}'.format(s_row), line.get('method_number'),
                             self.format_table)
            self.sheet.write('H{}'.format(s_row), line.get('depreciated_num'),
                             self.format_table)

            depreciated_remain = line.get('depreciated_remain')
            self.sheet.write('I{}'.format(s_row),
                             depreciated_remain >= 0 and
                             depreciated_remain or 0,
                             self.format_table)

            self.sheet.write('J{}'.format(s_row), line.get('value'),
                             self.format_table)
            self.sheet.write('K{}'.format(s_row), line.get('sum_debit'),
                             self.format_table)
            self.sheet.write(
                'L{}'.format(s_row), line.get('depreciated_value'),
                self.format_table)
            self.sheet.write('M{}'.format(s_row), '', self.format_table)
            depreciated_remain_value = \
                line.get('value') - line.get('sum_debit') - \
                line.get('depreciated_value')
            self.sheet.write('N{}'.format(s_row), depreciated_remain_value,
                             self.format_table)
            self.sheet.write('O{}'.format(s_row), line.get('dep_acc'),
                             self.format_table)
            self.sheet.write('P{}'.format(s_row), line.get('app_acc'),
                             self.format_table)

            stt += 1
            s_row += 1

        return True


AssetSummaryReportXlsx('report.asset_summary_xlsx_report',
                       'asset.summary.report.wizard')
