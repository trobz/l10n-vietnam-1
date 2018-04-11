# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from datetime import datetime

from openerp import fields
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT as DF
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT as DTF


class AccountStockLedgerReportXlsx(ReportXlsx):

    def generate_xlsx_report(self, workbook, data, objects):
        self.object = objects[0]
        self._define_formats(workbook)
        self.sheet = workbook.add_worksheet()
        self.setup_config()

        # generate header
        self.generate_company_info()

        # generate main content
        self.generate_report_title()
        self.generate_content_header()
        self.generate_main_content()

        self.generate_footer_content()

    def setup_config(self):
        self._set_default_format()

    def _define_formats(self, workbook):
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        format_config = {
            'font_name': 'Arial',
            'font_size': 10,
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
            'align': 'center',
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
        self.sheet.set_column('A:Z', None, self.format_default)
        self.sheet.set_row(2, 20)
        self.sheet.set_row(12, 20)
        self.sheet.set_row(13, 20)

        self.sheet.set_column('A:A', 10)
        self.sheet.set_column('B:B', 13)
        self.sheet.set_column('C:C', 10)
        self.sheet.set_column('D:D', 20)
        self.sheet.set_column('E:E', 10)
        self.sheet.set_column('F:F', 10)
        self.sheet.set_column('G:G', 10)
        self.sheet.set_column('H:H', 10)
        self.sheet.set_column('I:I', 10)
        self.sheet.set_column('J:J', 10)
        self.sheet.set_column('K:K', 10)
        self.sheet.set_column('L:L', 10)
        self.sheet.set_column('M:M', 10)

    def generate_company_info(self):
        company_info = self.object.env.user.company_id.get_company_info()

        self.sheet.merge_range(
            'A1:C1', u'Đơn vị báo cáo: ' + company_info['name'],
            self.format_company_info)

        self.sheet.merge_range(
            'A2:C2', u'Địa chỉ: ' + company_info['address'],
            self.format_company_info)

        self.sheet.merge_range(
            'A3:C3', u'MST: ', self.format_company_info)

    def generate_report_title(self):
        account_info = self.get_account_info()
        ledger_info = self.get_ledger_info()

        self.sheet.merge_range(
            'A4:M4', self.get_report_name(), self.format_report_title)

        self.sheet.merge_range(
            'A5:M5', u'Loại tài khoản: %s' % account_info['code'],
            self.format_center)

        self.sheet.merge_range(
            'A6:M6', u'Tên tài khoản: %s' % account_info['name'],
            self.format_center)

        self.sheet.merge_range(
            'A7:M7', u'Tên kho: %s' % ledger_info['location_name'],
            self.format_center)

        self.sheet.merge_range(
            'A8:M8', u'Mã nguyên liệu, vật liệu, công cụ, '
            u'dụng cu (sản phẩm, hàng hóa): %s' % ledger_info['product_code'],
            self.format_center)

        self.sheet.merge_range(
            'A9:M9', u'Tên nguyên liệu, vật liệu, công cụ, '
            u'dụng cu (sản phẩm, hàng hóa): %s' % ledger_info['product_name'],
            self.format_center)

        date_info = self.get_date_info()
        self.sheet.merge_range(
            'A10:M10', u'Từ %s đến %s' % (
                date_info['date_from'], date_info['date_to']),
            self.format_center)

        self.sheet.write(
            'M12', u'Đơn vị tính: %s' % self.get_uom(),
            self.format_uom)

    def get_uom(self):
        return self.object.product_tmpl_id.uom_id.name

    def get_ledger_info(self):
        return {
            'location_name': self.object.location_id.name,
            'product_code': self.object.product_id.default_code or
            self.object.product_tmpl_id.internal_ref_code,
            'product_name': self.object.product_id.name or
            self.object.product_tmpl_id.name,
        }

    def get_account_info(self):
        account = self.object.account_id
        return {
            'code': account.code,
            'name': account.name
        }

    def get_date_info(self):
        date_from = self.convert_date_format(self.object.date_from)
        date_to = self.convert_date_format(self.object.date_to)
        return {
            'date_from': date_from,
            'date_to': date_to
        }

    def convert_date_format(self, date_str, ctx_timestamp=False):
        rec = self.object
        fm = DF
        if ctx_timestamp:
            fm = DTF
        val = datetime.strptime(date_str, fm)
        if ctx_timestamp:
            val = fields.Datetime.context_timestamp(rec, val)
        return datetime.strftime(val, '%d/%m/%Y')
        try:
            rec = self.object
            fm = DF
            if ctx_timestamp:
                fm = DTF
            val = datetime.strptime(date_str, fm)
            if ctx_timestamp:
                val = fields.Datetime.context_timestamp(rec, val)
            return datetime.strftime(val, '%d/%m/%Y')

        except:
            return ''

    def get_report_name(self):
        return u'SỔ CHI TIẾT VẬT LIỆU, DỤNG CỤ (SẢN PHẨM, HÀNG HÓA)'

    def generate_content_header(self):
        self.sheet.merge_range(
            'A13:A14', u'Ngày tháng ghi sổ', self.format_table_bold_center)

        self.sheet.merge_range(
            'B13:C13', u'Chứng từ', self.format_table_bold_center)

        self.sheet.write(
            'B14', u'Số hiệu', self.format_table_bold_center)

        self.sheet.write(
            'C14', u'Ngày tháng', self.format_table_bold_center)

        self.sheet.merge_range(
            'D13:D14', u'Diễn giải', self.format_table_bold_center)

        self.sheet.merge_range(
            'E13:E14', u'TK đối ứng', self.format_table_bold_center)

        self.sheet.merge_range(
            'F13:F14', u'Đơn giá', self.format_table_bold_center)

        self.sheet.merge_range(
            'G13:H13', u'Nhập',
            self.format_table_bold_center)

        self.sheet.write(
            'G14', u'Số lượng', self.format_table_bold_center)

        self.sheet.write(
            'H14', u'Thành tiền', self.format_table_bold_center)

        self.sheet.merge_range(
            'I13:J13', u'Xuất', self.format_table_bold_center)

        self.sheet.write(
            'I14', u'Số lượng', self.format_table_bold_center)

        self.sheet.write(
            'J14', u'Thành tiền', self.format_table_bold_center)

        self.sheet.merge_range(
            'K13:L13', u'Tồn', self.format_table_bold_center)

        self.sheet.write(
            'K14', u'Số lượng', self.format_table_bold_center)

        self.sheet.write(
            'L14', u'Thành tiền', self.format_table_bold_center)

        self.sheet.merge_range(
            'M13:M14', u'Ghi chú', self.format_table_bold_center)

    def generate_main_content(self):
        self.row_position = 15
        start_position = self.row_position
        self.write_begining_inventory()
        main_contents = self.prepare_main_content()

        for line in main_contents:
            self.sheet.write(
                'A%s' % self.row_position, self.convert_date_format(
                    line['create_date'], True),
                self.format_table)

            self.sheet.write(
                'B%s' % self.row_position, line['ref'],
                self.format_table)

            self.sheet.write(
                'C%s' % self.row_position, line['date'],
                self.format_table)

            self.sheet.write(
                'D%s' % self.row_position, line['narration'],
                self.format_table)

            self.sheet.write(
                'E%s' % self.row_position, line['account'],
                self.format_table_center)

            self.sheet.write(
                'F%s' % self.row_position, line['unit_price'] or '',
                self.format_table_right)

            self.sheet.write(
                'G%s' % self.row_position, line['input_qty'] or '',
                self.format_table_right)

            self.sheet.write(
                'H%s' % self.row_position, line['input_amount'] or '',
                self.format_table_right)

            self.sheet.write(
                'I%s' % self.row_position, line['output_qty'] or '',
                self.format_table_right)

            self.sheet.write(
                'J%s' % self.row_position, line['output_amount'] or '',
                self.format_table_right)

            self.sheet.write(
                'K%s' % self.row_position, line['remain_qty'] or '',
                self.format_table_right)

            self.sheet.write(
                'L%s' % self.row_position, line['remain_amount'] or '',
                self.format_table_right)

            self.sheet.write(
                'M%s' % self.row_position, '',
                self.format_table_right)

            self.row_position += 1
        # ---------------------------------------------------------------------
        # Write SUM functions
        # ---------------------------------------------------------------------
        for i in range(0, 13):
            self.sheet.write(self.row_position - 1, i, '', self.format_table)
        self.sheet.set_row(self.row_position - 1, 20)
        self.sheet.write(
            'D%s' % self.row_position, u'Tổng cộng',
            self.format_table_bold_center)

        end_position = self.row_position - 1

        self.sheet.write_formula(
            'G%s' % self.row_position,
            '=SUM(G%s:G%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'H%s' % self.row_position,
            '=SUM(H%s:H%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'I%s' % self.row_position,
            '=SUM(I%s:I%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'J%s' % self.row_position,
            '=SUM(J%s:J%s)' % (start_position, end_position),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'K%s' % self.row_position,
            '=K%s' % (self.row_position - 1),
            self.format_table_bold_right)

        self.sheet.write_formula(
            'L%s' % self.row_position,
            '=L%s' % (self.row_position - 1),
            self.format_table_bold_right)

        self.row_position += 1
        return True

    def generate_footer_content(self):
        self.row_position += 2

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'Ngày ... tháng ... năm ...', self.format_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'Người ghi sổ', self.format_bold_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'Kế toán trưởng', self.format_bold_center)

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'Giám đốc', self.format_bold_center)
        self.row_position += 1

        self.sheet.merge_range(
            'A{0}:B{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'D{0}:E{0}'.format(self.row_position),
            u'(Ký, họ tên)', self.format_center)

        self.sheet.merge_range(
            'H{0}:I{0}'.format(self.row_position),
            u'(Ký, họ tên, đóng dấu)', self.format_center)

    def write_begining_inventory(self):
        self.begining_inventory = self.get_begining_inventory()
        for col in range(0, 13):
            self.sheet.write(self.row_position - 1, col, '', self.format_table)

        self.sheet.write(
            'D%s' % self.row_position, u'Số dư dầu kỳ',
            self.format_table_bold)

        quantity = self.begining_inventory['qty'] or ''
        self.sheet.write(
            'K%s' % self.row_position, quantity, self.format_table_right)

        amount = self.begining_inventory['amount'] or ''
        self.sheet.write(
            'L%s' % self.row_position, amount, self.format_table_right)
        self.row_position += 1

    def get_begining_inventory(self):
        params = self.get_params_for_query()
        sql = """
            SELECT COALESCE(SUM(sm.product_qty), 0 ) qty,
                COALESCE(SUM(sm.price_unit * sm.product_qty), 0) AS amount
                
            FROM 
            (
                SELECT sm.id, sm.product_id, price_unit, sm.date,
                    CASE
                    WHEN sm.location_id IN %(location_ids)s
                            OR pack.location_id IN %(location_ids)s
                        THEN -sm.product_qty
                        ELSE sm.product_qty
                    END,
                    
                    CASE
                    WHEN link.id IS NOT NULL
                        THEN pack.location_id
                        ELSE sm.location_id
                    END,
                
                    CASE
                    WHEN link.id IS NOT NULL
                    THEN pack.location_dest_id
                    ELSE sm.location_dest_id
                    END
                
                FROM stock_move sm
                    JOIN product_product pp
                        ON pp.id = sm.product_id
                    
                    JOIN product_template pt
                        ON pt.id = pp.product_tmpl_id
                    
                    LEFT JOIN stock_move_operation_link link
                    ON link.move_id = sm.id
                
                    LEFT JOIN stock_pack_operation pack
                    ON pack.id = link.operation_id
                
                WHERE sm.state='done'
                    {0}
                
            ) sm
                
            LEFT JOIN account_move am ON am.stock_move_id = sm.id
            LEFT JOIN account_move_line aml ON aml.move_id = am.id
                
            WHERE NOT (sm.location_id IN %(location_ids)s  
                AND sm.location_dest_id IN %(location_ids)s )
                AND (sm.location_id IN %(location_ids)s  
                    OR sm.location_dest_id IN %(location_ids)s )
                AND sm.date < '%(date_from)s'
                AND (
                    (aml.account_id = %(account_id)s 
                    AND am.state in %(state)s
                    AND aml.journal_id IN %(journal_ids)s
                    )
                    OR aml.account_id IS NULL
                )
                
        """

        sql = self.add_filter_to_query(sql)
        sql = sql % params

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()

        return data[0]

    def prepare_main_content(self):
        params = self.get_params_for_query()
        sql = """
            SELECT sm.id, date_trunc('seconds', sm.date) create_date,
                aml.ref,
                aml.date,
                am.narration,
                acc_du.code account,
                sm.price_unit unit_price,
                (CASE WHEN sm.product_qty > 0 THEN sm.product_qty ELSE 0 END) input_qty,
                (CASE WHEN sm.product_qty > 0 THEN abs(sm.product_qty * sm.price_unit) ELSE 0 END) input_amount,
                (CASE WHEN sm.product_qty < 0 THEN abs(sm.product_qty) ELSE 0 END) output_qty,
                (CASE WHEN sm.product_qty < 0 THEN abs(sm.product_qty * sm.price_unit) ELSE 0 END) output_amount
                    
            FROM 
            (
                SELECT sm.id, sm.product_id, price_unit, sm.date,
                    CASE
                    WHEN sm.location_id IN %(location_ids)s
                            OR pack.location_id IN %(location_ids)s
                        THEN -sm.product_qty
                        ELSE sm.product_qty
                    END,
                    
                    CASE
                    WHEN link.id IS NOT NULL
                        THEN pack.location_id
                        ELSE sm.location_id
                    END,
                
                    CASE
                    WHEN link.id IS NOT NULL
                    THEN pack.location_dest_id
                    ELSE sm.location_dest_id
                    END
                
                FROM stock_move sm
                    JOIN product_product pp
                        ON pp.id = sm.product_id
                    
                    JOIN product_template pt
                        ON pt.id = pp.product_tmpl_id
                    
                    LEFT JOIN LATERAL (
                        SELECT *
                        FROM stock_move_operation_link
                        WHERE move_id = sm.id
                        LIMIT 1
                    ) link ON TRUE
                    
                    LEFT JOIN stock_pack_operation pack
                            ON pack.id = link.operation_id
                WHERE sm.state='done'
                    {0}
            ) sm
                    
            LEFT JOIN account_move am ON am.stock_move_id = sm.id
            LEFT JOIN account_move_line aml ON aml.move_id = am.id
            
            LEFT JOIN LATERAL (
                SELECT *
                FROM account_move_line
                WHERE id = aml.counter_move_id OR counter_move_id = aml.id
            
            ) aml_counterpart ON TRUE
            
            LEFT JOIN account_account acc_du
                ON acc_du.id=aml_counterpart.account_id
                    
            WHERE NOT (sm.location_id IN %(location_ids)s  
                AND sm.location_dest_id IN %(location_ids)s )
                AND (sm.location_id IN %(location_ids)s  
                    OR sm.location_dest_id IN %(location_ids)s )
                AND sm.date BETWEEN '%(date_from)s' AND '%(date_to)s'
                AND (
                    (aml.account_id = %(account_id)s 
                    AND am.state in %(state)s
                    AND aml.journal_id IN %(journal_ids)s
                    )
                    OR aml.account_id IS NULL
                )
        """

        sql = self.add_filter_to_query(sql)
        sql = sql % params
        sql += """
            ORDER BY sm.date
        """

        self.env.cr.execute(sql)
        data = self.env.cr.dictfetchall()
        data = self.modify_main_data(data)
        return data

    def add_filter_to_query(self, sql):
        params = self.get_params_for_query()
        filter_product_sql = ''

        if params['product_id']:
            filter_product_sql += """
                AND sm.product_id = %(product_id)s
            """
        else:
            filter_product_sql += """
                AND pt.id = %(product_tmpl_id)s
            """

        return sql.format(filter_product_sql)

    def modify_main_data(self, data):
        first_line = True
        begining_inventory = self.begining_inventory
        remain_qty = remain_amount = 0.0

        for index, line in enumerate(data):
            if first_line:
                first_line = False
                remain_qty = line.get('input_qty', 0) - \
                    line.get('output_qty', 0) + \
                    begining_inventory['qty']

                remain_amount = line.get('input_amount', 0) - \
                    line.get('output_amount', 0) + \
                    begining_inventory['amount']

            else:
                previous_line = data[index - 1]
                remain_qty = line.get('input_qty', 0) - \
                    line.get('output_qty', 0) + \
                    previous_line.get('remain_qty', 0)

                remain_amount = line.get('input_amount', 0) - \
                    line.get('output_amount', 0) + \
                    previous_line.get('remain_amount', 0)

            line.update({
                'remain_qty': remain_qty,
                'remain_amount': remain_amount,
            })

        return data

    def get_params_for_query(self):
        location_id = self.object.location_id.id
        locations = self.env['stock.location'].search([
            ('id', 'child_of', location_id),
        ])
        location_ids = [-1, -1] + locations.ids

        journal_ids = self.object.journal_ids.ids
        if journal_ids:
            journal_ids += [-1, -1]
        else:
            journal_ids = [-1, -1]

        params = {
            'account_id': self.object.account_id.id,
            'product_id': self.object.product_id.id,
            'product_tmpl_id': self.object.product_tmpl_id.id,
            'location_ids': tuple(location_ids),
            'date_from': self.object.date_from_dt,
            'date_to': self.object.date_to_dt,
            'journal_ids': tuple(journal_ids)
        }

        if self.object.target_move == 'posted':
            params['state'] = ('posted', '1')

        elif self.object.target_move == 'all':
            params['state'] = ('posted', 'draft')

        return params


AccountStockLedgerReportXlsx(
    'report.stock_ledger_report', 'account.stock.ledger.wizard')
