# -*- coding: utf-8 -*-
# Copyright 2009-2018 Trobz (http://trobz.com)
# License AGPL-3.0 or later (https://www.gnu.org/licenses/agpl).

from openerp import api, models
import logging


class post_object_report_account_vas(models.AbstractModel):
    _name = "post.object.report.account.vas"
    _description = "Set up Data for VAS reports"

    @api.model
    def start(self):
        # Functions run one time
        PostFunction = self.env['post.function.one.time']
        PostFunction.run_post_object_one_time(
            'post.object.report.account.vas',
            [
                'set_balance_sheet_config_data_one_time',
                'set_profit_and_loss_config_data_one_time',
                'set_cash_flow_indirect_config_one_time'
            ]
        )

        # Functions run all times
        # HERE
        return True

    @api.model
    def set_balance_sheet_config_data_one_time(self):
        logging.info('START -- create balance sheet config data ...')
        LINES_CONFIG_DATA = [
            {
                'item': 'Tiền',
                'code': 111,
                'is_parenthesis': False,
                'is_inverted_result': False,
                'config_line_ids': [
                    (0, 0, {
                        'code': '111',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '112',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '113',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Các khoản tương đương tiền',
                'code': 112,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1281',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Chứng khoán kinh doanh',
                'code': 121,
                'config_line_ids': [
                    (0, 0, {
                        'code': '121',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng giảm giá chứng khoán kinh doanh',
                'code': 122,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2291',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Đầu tư nắm giữ đến ngày đáo hạn',
                'code': 123,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1282',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '1288',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Phải thu ngắn hạn của khách hàng',
                'code': 131,
                'config_line_ids': [
                    (0, 0, {
                        'code': '131',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Trả trước cho người bán ngắn hạn',
                'code': 132,
                'config_line_ids': [
                    (0, 0, {
                        'code': '331',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải thu nội bộ ngắn hạn',
                'code': 133,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1362',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '1363',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '1368',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': (
                    'Phải thu theo tiến độ kế hoạch '
                    'hợp đồng xây dựng'
                ),
                'code': 134,
                'config_line_ids': [
                    (0, 0, {
                        'code': '337',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải thu về cho vay ngắn hạn',
                'code': 135,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1283',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải thu ngắn hạn khác',
                'code': 136,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1385',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '1388',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '334',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '338',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '141',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '244',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng phải thu ngắn hạn khó đòi',
                'code': 137,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2293',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Tài sản thiếu chờ xử lý',
                'code': 139,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1381',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Hàng tồn kho',
                'code': 141,
                'config_line_ids': [
                    (0, 0, {
                        'code': '151',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '152',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '153',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '154',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '155',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '156',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '157',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '158',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng giảm giá hàng tồn kho (*)',
                'code': 149,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2294',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Chi phí trả trước ngắn hạn',
                'code': 151,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2421',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Thuế GTGT được khấu trừ',
                'code': 152,
                'config_line_ids': [
                    (0, 0, {
                        'code': '133',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': ' Thuế và các khoản khác phải thu Nhà nước',
                'code': 153,
                'config_line_ids': [
                    (0, 0, {
                        'code': '333',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Giao dịch mua bán lại trái phiếu Chính phủ',
                'code': 154,
                'config_line_ids': [
                    (0, 0, {
                        'code': '171',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Tài sản ngắn hạn khác',
                'code': 155,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2288',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Phải thu dài hạn của khách hàng',
                'code': 211,
                'config_line_ids': []
            },

            {
                'item': 'Trả trước cho người bán dài hạn',
                'code': 212,
                'config_line_ids': []
            },

            {
                'item': 'Vốn kinh doanh ở đơn vị trực thuộc',
                'code': 213,
                'config_line_ids': [
                    (0, 0, {
                        'code': '1361',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Phải thu nội bộ dài hạn',
                'code': 214,
                'config_line_ids': []
            },

            {
                'item': 'Phải thu về cho vay dài hạn',
                'code': 215,
                'config_line_ids': []
            },

            {
                'item': 'Phải thu dài hạn khác',
                'code': 216,
                'config_line_ids': []
            },

            {
                'item': 'Dự phòng phải thu dài hạn khó đòi',
                'code': 219,
                'is_parenthesis': True,
                'config_line_ids': []
            },
            {
                'item': 'Nguyên giá',
                'code': 222,
                'config_line_ids': [
                    (0, 0, {
                        'code': '211',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Giá trị hao mòn lũy kế',
                'code': 223,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2141',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Nguyên giá',
                'code': 225,
                'config_line_ids': [
                    (0, 0, {
                        'code': '212',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Giá trị hao mòn luỹ kế (*)',
                'code': 226,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2142',
                        'is_credit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Nguyên giá',
                'code': 228,
                'config_line_ids': [
                    (0, 0, {
                        'code': '213',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Giá trị hao mòn luỹ kế (*)',
                'code': 229,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2143',
                        'is_credit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Nguyên giá',
                'code': 231,
                'config_line_ids': [
                    (0, 0, {
                        'code': '217',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Giá trị hao mòn luỹ kế (*)',
                'code': 232,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2147',
                        'is_credit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Chi phí sản xuất, kinh doanh dở dang dài hạn',
                'code': 241,
                'config_line_ids': []
            },
            {
                'item': 'Chi phí xây dựng cơ bản dở dang',
                'code': 242,
                'config_line_ids': [
                    (0, 0, {
                        'code': '241',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Đầu tư vào công ty con ',
                'code': 251,
                'config_line_ids': [
                    (0, 0, {
                        'code': '221',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Đầu tư vào công ty liên kết, liên doanh',
                'code': 252,
                'config_line_ids': [
                    (0, 0, {
                        'code': '222',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Đầu tư góp vốn vào đơn vị khác',
                'code': 253,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2281',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng đầu tư tài chính dài hạn',
                'code': 254,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2292',
                        'is_credit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Đầu tư nắm giữ đến ngày đáo hạn',
                'code': 255,
                'config_line_ids': []
            },

            # ===================================================
            # Tài sản dài hạn khác
            # ===================================================
            {
                'item': 'Chi phí trả trước dài hạn',
                'code': 261,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2422',
                        'is_debit_balance': True,
                    }),
                    (0, 0, {
                        'code': '2423',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Tài sản thuế thu nhập hoãn lại',
                'code': 262,
                'config_line_ids': [
                    (0, 0, {
                        'code': '243',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Thiết bị, vật tư, phụ tùng thay thế dài hạn',
                'code': 263,
                'config_line_ids': []
            },
            {
                'item': 'Tài sản dài hạn khác',
                'code': 268,
                'config_line_ids': [
                    (0, 0, {
                        'code': '2288',
                        'is_debit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải trả người bán ngắn hạn',
                'code': 311,
                'is_inverted_result': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '331',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Người mua trả tiền trước ngắn hạn',
                'code': 312,
                'is_inverted_result': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '131',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Thuế và các khoản phải nộp Nhà nước',
                'code': 313,
                'config_line_ids': [
                    (0, 0, {
                        'code': '333',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải trả người lao động',
                'code': 314,
                'config_line_ids': [
                    (0, 0, {
                        'code': '334',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Chi phí phải trả ngắn hạn',
                'code': 315,
                'config_line_ids': [
                    (0, 0, {
                        'code': '335',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải trả nội bộ ngắn hạn',
                'code': 316,
                'config_line_ids': [
                    (0, 0, {
                        'code': '3362',
                        'is_credit_balance': True,
                    }),
                    (0, 0, {
                        'code': '3363',
                        'is_credit_balance': True,
                    }),
                    (0, 0, {
                        'code': '3368',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': (
                    'Phải trả theo tiến độ '
                    'kế hoạch hợp đồng xây dựng'
                ),
                'code': 317,
                'config_line_ids': [
                    (0, 0, {
                        'code': '337',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Doanh thu chưa thực hiện ngắn hạn',
                'code': 318,
                'config_line_ids': [
                    (0, 0, {
                        'code': '3387',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải trả ngắn hạn khác',
                'code': 319,
                'config_line_ids': [
                    (0, 0, {
                        'code': '338',
                        'is_credit_balance': True,
                    }),
                    (0, 0, {
                        'code': '138',
                        'is_credit_balance': True,
                    }),
                    (0, 0, {
                        'code': '3441',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Vay và nợ thuê tài chính ngắn hạn',
                'code': 320,
                'config_line_ids': [
                    (0, 0, {
                        'code': '341',
                        'is_credit_balance': True,
                    }),
                    (0, 0, {
                        'code': '34311',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng phải trả ngắn hạn',
                'code': 321,
                'config_line_ids': [
                    (0, 0, {
                        'code': '352',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Quỹ khen thưởng, phúc lợi',
                'code': 322,
                'config_line_ids': [
                    (0, 0, {
                        'code': '353',
                        'is_debit_balance': True,
                        'is_credit_balance': True,
                        'is_inverted': True,
                    }),
                ]
            },

            # ===================================================
            # ========== Nợ dài hạn ===============
            # ===================================================
            # {
            #     'item': 'Người mua trả tiền trước dài hạn',
            #     'code': 332,
            #     'config_line_ids': []
            # },
            {
                'item': 'Chi phí phải trả dài hạn',
                'code': 333,
                'config_line_ids': []
            },
            {
                'item': 'Phải trả nội bộ về vốn kinh doanh',
                'code': 334,
                'config_line_ids': [
                    (0, 0, {
                        'code': '3361',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Phải trả nội bộ dài hạn',
                'code': 335,
                'config_line_ids': []
            },
            {
                'item': 'Doanh thu chưa thực hiện dài hạn',
                'code': 336,
                'config_line_ids': []
            },
            {
                'item': 'Phải trả dài hạn khác',
                'code': 337,
                'is_inverted_result': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '3442',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Vay và nợ thuê tài chính dài hạn',
                'code': 338,
                'config_line_ids': []
            },
            {
                'item': 'Trái phiếu chuyển đổi',
                'code': 339,
                'config_line_ids': [
                    (0, 0, {
                        'code': '3432',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Cổ phiếu ưu đãi',
                'code': 340,
                'config_line_ids': [
                    (0, 0, {
                        'code': '41112',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Thuế thu nhập hoãn lại phải trả',
                'code': 341,
                'config_line_ids': [
                    (0, 0, {
                        'code': '347',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Dự phòng phải trả dài hạn',
                'code': 342,
                'config_line_ids': [
                    (0, 0, {
                        'code': '352',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Quỹ phát triển khoa học và công nghệ',
                'code': 343,
                'config_line_ids': [
                    (0, 0, {
                        'code': '356',
                        'is_credit_balance': True,
                    }),
                ]
            },

            # ===================================================
            # =================== Vốn Chủ Sở Hữu ===================
            # ===================================================
            {
                'item': 'Vốn góp của chủ sở hữu',
                'code': '411',
                'config_line_ids': [
                    (0, 0, {
                        'code': '4111',
                    }),
                ]
            },
            {
                'item': 'Cổ phiếu phổ thông có quyền biểu quyết',
                'code': '411a',
                'config_line_ids': [
                    (0, 0, {
                        'code': '41111',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Cổ phiếu ưu đãi',
                'code': '411b',
                'config_line_ids': [
                    (0, 0, {
                        'code': '41112',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Thặng dư vốn cổ phần',
                'code': 412,
                'config_line_ids': [
                    (0, 0, {
                        'code': '4112',
                        'is_debit_balance': True,
                        'is_credit_balance': True,
                        'is_inverted': True,
                    }),
                ]
            },
            {
                'item': 'Quyền chọn chuyển đổi trái phiếu',
                'code': 413,
                'config_line_ids': [
                    (0, 0, {
                        'code': '4113',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Vốn khác của chủ sở hữu',
                'code': 414,
                'config_line_ids': [
                    (0, 0, {
                        'code': '4118',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Cổ phiếu quỹ',
                'code': 415,
                'is_parenthesis': True,
                'config_line_ids': [
                    (0, 0, {
                        'code': '419',
                        'is_debit_balance': True,
                    }),
                ]
            },

            {
                'item': 'Chênh lệch đánh giá lại tài sản',
                'code': 416,
                'config_line_ids': [
                    (0, 0, {
                        'code': '412',
                        'is_debit_balance': True,
                        'is_credit_balance': True,
                        'is_inverted': True,
                    }),
                ]
            },
            {
                'item': 'Chênh lệch tỷ giá hối đoái',
                'code': 417,
                'config_line_ids': [
                    (0, 0, {
                        'code': '413',
                        'is_debit_balance': True,
                        'is_credit_balance': True,
                        'is_inverted': True,
                    }),
                ]
            },
            {
                'item': 'Quỹ đầu tư phát triển',
                'code': 418,
                'config_line_ids': [
                    (0, 0, {
                        'code': '414',
                        'is_credit_balance': True,
                    }),
                ]
            },
            {
                'item': 'Quỹ hỗ trợ sắp xếp doanh nghiệp',
                'code': 419,
                'config_line_ids': [
                    (0, 0, {
                        'code': '417',
                        'is_credit_balance': True,
                    }),
                ]
            },
            # {
            #     'item': 'Quỹ khác thuộc vốn chủ sở hữu',
            #     'code': 420,
            #     'account_ids': [418],
            #     'config_line_ids': [
            #         (0, 0, {
            #             'code': '418',
            #             'is_credit_balance': True,
            #         }),
            #     ]
            # },
            # {
            #     'item': 'Lợi nhuận sau thuế chưa phân phối '
            #     'lũy kế đến cuối kỳ trước',
            #     'code': "421a",
            #     'account_ids': [4211],
            #     'is_debit': True, 'is_credit': True,
            #     'is_parenthesis': False, 'is_inverted_result': True,
            #     'config_line_ids': [
            #         (0, 0, {
            #             'code': '4211',
            #             'is_debit_balance': True,
            #             'is_credit_balance': True,
            #             'is_inverted': True,
            #         }),
            #     ]
            # },
            # {
            #     'item': 'Lợi nhuận sau thuế chưa phân phối kỳ này',
            #     'code': "421b",
            #     'config_line_ids': [
            #         (0, 0, {
            #             'code': '4212',
            #             'is_debit_balance': True,
            #             'is_credit_balance': True,
            #             'is_inverted': True,
            #         }),
            #     ]
            # },
            # {
            #     'item': 'Nguồn vốn đầu tư xây dựng cơ bản',
            #     'code': 422,
            #     'config_line_ids': [
            #         (0, 0, {
            #             'code': '441',
            #             'is_credit_balance': True,
            #         }),
            #     ]
            # },
            # {
            #     'item': 'Nguồn kinh phí',
            #     'code': 431,
            #     'is_inverted_result': True,
            #     'config_line_ids': [
            #         (0, 0, {
            #             'code': '1611',
            #             'is_debit_balance': True,
            #         }),
            #         (0, 0, {
            #             'code': '1612',
            #             'is_debit_balance': True,
            #         }),
            #         (0, 0, {
            #             'code': '4611',
            #             'is_credit_balance': True,
            #             'operator': 'minus'
            #         }),
            #         (0, 0, {
            #             'code': '4612',
            #             'is_credit_balance': True,
            #             'operator': 'minus'
            #         }),
            #     ]
            # },
            # {'item': 'Nguồn kinh phí đã hình thành TSCĐ',
            #  'code': 432,
            #  'config_line_ids': [
            #      (0, 0, {
            #          'code': '466',
            #          'is_credit_balance': True,
            #      }),
            #  ]
            #  }
        ]

        # create configuration data for account_balance_sheet report
        balance_sheet_config_env = self.env['balance.sheet.config']
        for config in LINES_CONFIG_DATA:
            balance_sheet_config_env.create(config)
        logging.info('END -- create balance sheet config data ...')
        return True

    @api.model
    def set_profit_and_loss_config_data_one_time(self):
        logging.info(
            'START -- create profit and loss config data ...'
        )
        LINES_CONFIG_DATA = [
            {'item': u'1. Doanh thu bán hàng và cung cấp dịch vụ',
             'code': '01',
             'account_ids': [511],
             'counterpart_account_ids': [],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': 'VI.25'},
            {'item': u'2. Các khoản giảm trừ doanh thu',
             'code': '02',
             'account_ids': [511],
             'counterpart_account_ids': [521],
             'is_debit': True,
             'is_credit': False,
             'exception': False,
             'note': ''},
            {'item': u'4. Giá vốn hàng bán',
             'code': '11',
             'account_ids': [632],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note':'VI.27'},
            {'item': u'6. Doanh thu hoạt động tài chính',
             'code': '21',
             'account_ids': [515],
             'counterpart_account_ids': [911],
             'is_debit': True,
             'is_credit': False,
             'exception': False,
             'note': 'VI.26'},
            {'item': u'7. Chi phí tài chính',
             'code': '22',
             'account_ids': [635],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': ''},
            {'item': u'   - Trong đó: Chi phí lãi vay',
             'code': '23',
             'account_ids': [6352],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': ''},
            {'item': u'8. Chi phí bán hàng',
             'code': '25',
             'account_ids': [641],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': ''},
            {'item': u'9. Chi phí quản lý doanh nghiệp',
             'code': '26',
             'account_ids': [642],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': ''},
            {'item': u'11. Thu nhập khác',
             'code': '31',
             'account_ids': [711],
             'counterpart_account_ids': [911],
             'is_debit': True,
             'is_credit': False,
             'exception': False,
             'note': ''},
            {'item': u'12. Chi phí khác',
             'code': '32',
             'account_ids': [811],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': True,
             'exception': False,
             'note': ''},
            {'item': u'15. Chi phí thuế thu nhập doanh nghiệp hiện hành',
             'code': '51',
             'account_ids': [8211],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': False,
             'exception': True,
             'note': 'VI.30'},
            {'item': u'16. Chi phí thuế thu nhập doanh nghiệp hoãn lại',
             'code': '52',
             'account_ids': [8212],
             'counterpart_account_ids': [911],
             'is_debit': False,
             'is_credit': False,
             'exception': True,
             'note': 'VI.30'}
        ]

        # create configuration data for account_balance_sheet report
        profit_and_loss_config_env = self.env['profit.and.loss.config']
        account_env = self.env['account.account']
        for line in LINES_CONFIG_DATA:
            account_ids = []
            counterpart_account_ids = []
            for acc in line.get('account_ids', []):
                account_ids += account_env.search(
                    [('code', '=like', '%s%%' % acc)]).ids
            for acc in line.get('counterpart_account_ids', []):
                counterpart_account_ids += account_env.search(
                    [('code', '=like', '%s%%' % acc)]).ids
            account_ids = [(6, 0, account_ids)]
            counterpart_account_ids = [(6, 0, counterpart_account_ids)]
            line.update({
                'account_ids': account_ids,
                'counterpart_account_ids': counterpart_account_ids
            })
            profit_and_loss_config_env.create(line)
        logging.info('END -- create profit and loss config data ...')
        return True

    @api.model
    def set_cash_flow_indirect_config_one_time(self):
        logging.info(
            'START -- create indirect cashflow config data ...'
        )
        LINES_CONFIG_DATA = [
            {
                'item': 'Lợi nhuận trước thuế',
                'code': '01',
                'child_ids': [
                    {
                        'item': 'Lợi nhuận trước thuế',
                        'code': '01',
                        'dr_account_ids': [911],
                        'cr_account_ids': [4212],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                    {
                        'item': 'Lợi nhuận trước thuế',
                        'code': '01',
                        'dr_account_ids': [4212],
                        'cr_account_ids': [911],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    }
                ]
            },


            {
                'item': 'Khấu hao TSCĐ',
                'code': '02',
                'child_ids': [
                    {
                        'item': (
                            'Số khấu hao TSCĐ đó trích vào '
                            'chi phí sản xuất, '
                            'kinh doanh trong kỳ báo cáo'
                        ),
                        'code': '02',
                        'dr_account_ids': [6274, 6414, 6424],
                        'cr_account_ids': [214],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    }
                ]
            },


            {
                'item': 'Các khoản dự phòng',
                'code': '03',
                'child_ids': [
                    {
                        'item': 'Dự phòng giảm giá đầu tư ngắn hạn, dài hạn',
                        'code': '03',
                        'dr_account_ids': [635],
                        'cr_account_ids': [129, 229],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''},

                    {
                        'item': 'Dự phòng phải thu khó đòi',
                        'code': '03',
                        'dr_account_ids': [642],
                        'cr_account_ids': [139],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Dự phòng giảm giá hàng tồn kho',
                        'code': '03',
                        'dr_account_ids': [632],
                        'cr_account_ids': [159],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },


                    {
                        'item': (
                            'Hoàn nhập dự phòng giảm giá đầu tư ngắn hạn, '
                            'dài hạn'
                        ),
                        'code': '03',
                        'dr_account_ids': [129, 229],
                        'cr_account_ids': [515],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Hoàn nhập dự phòng phải thu khó đòi',
                        'code': '03',
                        'dr_account_ids': [139],
                        'cr_account_ids': [711],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Hoàn nhập dự phòng giảm giá hàng tồn kho',
                        'code': '03',
                        'dr_account_ids': [159],
                        'cr_account_ids': [711],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Lãi, lỗ chênh lệch tỷ giá hối đoái chưa thực hiện',
                'code': '04',
                'child_ids': [

                    {
                        'item': 'Lời chênh lệch tỷ giá hối đoái',
                        'code': '04',
                        'dr_account_ids': [413],
                        'cr_account_ids': [515],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''},

                    {
                        'item': 'Lỗ chênh lệch tỷ giá hối đoái',
                        'code': '04',
                        'dr_account_ids': [635],
                        'cr_account_ids': [413],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Lãi, lỗ từ hoạt động đầu tư',
                'code': '05',
                'child_ids': [

                    {
                        'item': 'Phần thu thanh lý, nhượng bán TSCĐ 1',
                        'code': '05',
                        'dr_account_ids': [111, 112, 113, 131, 138],
                        'cr_account_ids': [7111, 5151, 33311],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Phần thu thanh lý, nhượng bán TSCĐ 2',
                        'code': '05',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [1311],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Phần chi thanh lý, nhượng bán TSCĐ',
                        'code': '05',
                        'dr_account_ids': [811, 6351, 1331],
                        'cr_account_ids': [111, 112, 113, 331, 338],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Lợi nhuận được chia từ khoản '
                            'đầu tư vốn vào đơn vị khác'
                        ),
                        'code': '05',
                        'dr_account_ids': [111, 112, 138, 222],
                        'cr_account_ids': [5153],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Hoàn nhập dự phòng phải thu khó đòi',
                        'code': '05',
                        'dr_account_ids': [111, 112, 121, 221],
                        'cr_account_ids': [5154],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    }
                ]
            },


            {
                'item': 'Chi phí lãi vay',
                'code': '06',
                'child_ids': [
                    {
                        'item': (
                            'Chi phí lãi vay phát sinh và đó ghi nhận vào '
                            'kết quả kinh doanh trong kỳ'
                        ),
                        'code': '06',
                        'dr_account_ids': [6352],
                        'cr_account_ids': [111, 112, 341, 311],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    }
                ]
            },


            {
                'item': 'Tăng, giảm các khoản phải thu',
                'code': '09',
                'child_ids': [
                    {
                        'item': (
                            'Chênh lệch số dư cuối kỳ (SDCK) và số dư '
                            'đầu kỳ (SDDK) phải thu khách hàng (mã131)'
                        ),
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [131],
                        'is_positive_difference': False,
                        'python_formular': ''},

                    {
                        'item': '+ Phải thu liên quan đến thanh lý TSCĐ',
                        'code': '09',
                        'dr_account_ids': [131],
                        'cr_account_ids': [7112, 5152, 33311],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': '- Thu tiền liên quan đến thanh lý TSCĐ',
                        'code': '09',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [1312],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK trả trước cho '
                            'người bán (mã 132)'
                        ),
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [331],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK phải thu '
                            'nội bộ (mã 134)'
                        ),
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [136],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK phải thu khác (mã 138)'
                        ),
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [138],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Phải thu liên quan đến thanh lý TSCĐ',
                        'code': '09',
                        'dr_account_ids': [1388],
                        'cr_account_ids': [7113, 5159, 33311],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Thu tiền liên quan đến thanh lý TSCĐ',
                        'code': '09',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [1388],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK Thuế GTGT '
                            'được khấu trừ (mã 133)'
                        ),
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [133],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Chênh lệch SDCK-SDDK tạm ứng (mã 151)',
                        'code': '09',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [141],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tăng, giảm hàng tồn kho',
                'code': '10',
                'child_ids': [
                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK hàng mua '
                            'đang đi trên đường (mã 151)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [151],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK nguyên vật'
                            ' liệu tồn kho (mã 152)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [152],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK cụng cụ, '
                            'dụng cụ trong kho (mã 153)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [153],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK chi phí SXKDDD (mã 154)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [154],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': 'Chênh lệch SDCK-SDDK thành phẩm (mã 155)',
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [155],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK hàng hóa'
                            ' tồn kho (mã 156)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [156],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK hàng gửi'
                            ' đi bán (mã 157)'
                        ),
                        'code': '10',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [157],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': (
                    'Tăng, giảm các khoản phải trả '
                    '(Không kể lãi vay phải trả, '
                    'thuế thu nhập doanh nghiệp phải nộp)'
                ),
                'code': '11',
                'child_ids': [

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK phải trả '
                            'cho người bán (mã 313)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [331],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK người mua'
                            ' trả tiền trước (mã 314)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [131],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK Thuế và các '
                            'khỏan phải nộp nhà nước (mã 315)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [333],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': 'Thuế thu nhập doanh nghiệp phải trả',
                        'code': '11',
                        'dr_account_ids': [421],
                        'cr_account_ids': [3334],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chi nộp thuế TNDN (không phân biệt số thuế '
                            'TNDN đó nộp của kỳ này, số thuế TNDN cũng '
                            'nợ từ các kỳ trước đó nộp trong kỳ này '
                            'và số thuế TNDN nộp trước (nếu có))'
                        ),
                        'code': '11',
                        'dr_account_ids': [3334],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK Phải trả '
                            'công nhân viên (mã 316)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [334],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK chi phí '
                            'phải trả (mã 331)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [335],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chi phí lãi vay phát sinh và đó ghi '
                            'nhận vào kết quả kinh doanh trong kỳ'
                        ),
                        'code': '11',
                        'dr_account_ids': [6353],
                        'cr_account_ids': [111, 112, 341, 311],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chi trả lãi vay (không phân biệt trả cho kỳ '
                            'trước, trả trong kỳ và trả trước lãi vay)'
                        ),
                        'code': '11',
                        'dr_account_ids': [336, 6354],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK phải '
                            'trả nội bộ (mã 317)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [336],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK các khoản phải trả, '
                            'phải nộp khác (mã 318)'
                        ),
                        'code': '11',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [338],
                        'is_positive_difference': True,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tăng, giảm chi phí trả trước.',
                'code': '12',
                'child_ids': [
                    {
                        'item': (
                            'Chênh lệch SDCK - SDDK chi '
                            'phí trả trước (mã 152)'
                        ),
                        'code': '12',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [142],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chênh lệch SDCK-SDDK chi phí '
                            'trả trước dài hạn (mã 241)'
                        ),
                        'code': '12',
                        'dr_account_ids': [],
                        'cr_account_ids': [],
                        'is_inverted_result': False,
                        'is_formula_active': True,
                        'account_ids': [242],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tiền lãi vay đã trả',
                'code': '13',
                'child_ids': [

                    {
                        'item': (
                            'Chi trả lời vay (không phân biệt trả cho '
                            'kỳ trước, trả trong kỳ và trả trước lời vay)'
                        ),
                        'code': '13',
                        'dr_account_ids': [335, 6355],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': '''Thuế thu nhập doanh nghiệp đó nộp''',
                'code': '14',
                'child_ids': [

                    {
                        'item': (
                            'Chi nộp thuế TNDN (khụng phân biệt số '
                            'thuế TNDN đó nộp của kỳ này, số thuế TNDN cũng '
                            'nợ từ các kỳ trước đó nộp trong kỳ này '
                            'và số thuế TNDN nộp trước (nếu có))'
                        ),
                        'code': '14',
                        'dr_account_ids': [3334],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền khu khác từ hoạt động kinh doanh',
                'code': '15',
                'child_ids': [

                    {
                        'item': (
                            'Tiền thu do nhận ký quỹ, ký cược + '
                            'Thu hồi các khoản đưa đi ký quỹ, ký cược + '
                            'Tiền từ các tổ chức cỏ nhõn bờn ngoài thưởng, '
                            'hỗ trợ ghi tăng quỹ doanh nghiệp'
                        ),
                        'code': '15',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [344, 144, 244, 431],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền chi khác cho hoạt động kinh doanh',
                'code': '16',
                'child_ids': [
                    {
                        'item': (
                            'Tiền trả các khoản nhận ký cược, ký quỹ + '
                            'Tiền chi đưa đi ký quỹ, ký cược + '
                            'Tiền chi trực tiếp từ quỹ khen thưởng, phúc '
                            'lợi và các quỹ khác'
                        ),
                        'code': '16',

                        'dr_account_ids': [344, 144, 244, 431],
                        'cr_account_ids': [111, 112],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': (
                    'Tiền chi để mua sắm, xây dựng TSCĐ và '
                    'các tài sản dài hạn khác'
                ),
                'code': '21',
                'child_ids': [

                    {
                        'item': (
                            'Chi tiền mua sắm TSCĐ + ứng tiền cho nhà thầu, '
                            'nhà cung cấp đầu tư TSCĐ và các TS dài hạn khác'
                        ),
                        'code': '21',
                        'dr_account_ids': [211, 331],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Chi phí xây dựng cơ bản phát sinh bằng tiền'
                        ),
                        'code': '21',
                        'dr_account_ids': [241, 1332],
                        'cr_account_ids': [111, 112],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': (
                    'Tiền thu từ thanh lý, nhượng bán TSCĐ và '
                    'các tài sản dài hạn khác'
                ),
                'code': '22',
                'child_ids': [

                    {
                        'item': '''1 - Phần thu thanh lý, nhượng bán TSCĐ''',
                        'code': '22',
                        'dr_account_ids': [111, 112, 113],
                        'cr_account_ids': [7113, 5155, 33311],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': '''1 - Phần thu thanh lý, nhượng bán TSCĐ''',
                        'code': '22',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [1311, 138],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': '''2 - Phần chi thanh lý, nhượng bán TSCĐ''',
                        'code': '22',
                        'dr_account_ids': [811, 6356, 1331, 331, 338],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền chi mua các công cụ nợ của đơn vị khác',
                'code': '23',
                'child_ids': [
                    {
                        'item': (
                            'Chi cho doanh nghiệp khác vay + '
                            'Chi mua trái phiếu, tín phiếu, kỳ phiếu'
                        ),
                        'code': '23',
                        'dr_account_ids': [128, 228, 121, 221],
                        'cr_account_ids': [111, 112],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': (
                    'Tiền thu từ thanh lý các công cụ nợ của đơn vị khác'
                ),
                'code': '24',
                'child_ids': [
                    {
                        'item': (
                            'Thanh tóan nợ gốc trái phiếu, tín phiếu, '
                            'kỳ phiếu + Thu hồi nợ gốc các doanh nghiệp '
                            'khác vay + Tiền thu do bán lại trái phiếu, '
                            'tín phiếu, kỳ phiếu'
                        ),
                        'code': '24',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [1281, 228, 121, 221, 222],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền chi đầu tư góp vốn vào đơn vị khác',
                'code': '25',
                'child_ids': [
                    {
                        'item': '''Góp vốn vào các doanh nghiệp khác''',
                        'code': '25',
                        'dr_account_ids': [221, 222, 128, 228],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền thu hồi đầu tư góp vốn vào đơn vị khác',
                'code': '26',
                'child_ids': [
                    {
                        'item': 'Thu hồi gúp vốn vào các doanh nghiệp khác',
                        'code': '26',
                        'dr_account_ids': [111, 112, 113],
                        'cr_account_ids': [221, 222, 128, 228],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': 'Tiền thu cổ tức và lợi nhuận được chia',
                'code': '27',
                'child_ids': [
                    {
                        'item': (
                            'Thu lời tiền gửi ngõn hàng + Thu cổ tức, '
                            'lợi nhuận được chia (nếu có)'
                        ),
                        'code': '27',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [515, 33311],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': (
                    'Tiền thu từ phát hành cổ phiếu, '
                    'trái phiếu, nhận vốn góp của chủ sở hữu'
                ),
                'code': '31',
                'child_ids': [
                    {
                        'item': (
                            'Tiền thu do nhận vốn trực tiếp từ NSNN hoặc '
                            'do các chủ sở hữu gúp vốn + Nhận cấp phát vốn '
                            'đầu tư XDCB'
                        ),
                        'code': '31',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [411, 414],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''}
                ]
            },


            {
                'item': (
                    'Tiền chi trả vốn góp cho các chủ sở hữu, '
                    'mua lại cổ phiếu của công ty đã phát hành'
                ),
                'code': '32',
                'child_ids': [
                    {
                        'item': (
                            'Hoàn trả vốn trực tiếp cho NSNN '
                            'hoặc chủ sở hữu'
                        ),
                        'code': '32',
                        'dr_account_ids': [411],
                        'cr_account_ids': [111, 112],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tiền vay ngắn hạn, dài hạn nhận được',
                'code': '33',
                'child_ids': [
                    {
                        'item': (
                            'Hoàn trả vốn trực tiếp cho NSNN '
                            'hoặc chủ sở hữu'
                        ),
                        'code': '33',
                        'dr_account_ids': [111, 112],
                        'cr_account_ids': [311, 341],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tiền chi trả nợ gốc vay',
                'code': '34',
                'child_ids': [
                    {
                        'item': ('Trả tiền vay ngắn hạn cho các '
                                 'tổ chức tín dụng'
                                 ),
                        'code': '34',
                        'dr_account_ids': [311],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },

                    {
                        'item': (
                            'Trả tiền vay dài hạn cho các '
                            'tổ chức tín dụng'
                        ),
                        'code': '34',
                        'dr_account_ids': [341, 315],
                        'cr_account_ids': [111, 112],
                        'is_inverted_result': False,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Tiền chi trả nợ thuê tài chính',
                'code': '35',
                'child_ids': [
                    {
                        'item': '''Tiền chi trả nợ thuê tài chánh''',
                        'code': '35',
                        'dr_account_ids': [315, 342],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },


            {
                'item': 'Cổ tức, lợi nhuận đã trả cho chủ sở hữu',
                'code': '36',
                'child_ids': [
                    {
                        'item': (
                            'Cổ tức, lợi nhuận đó trả cho chủ sở hữu'
                        ),
                        'code': '36',
                        'dr_account_ids': [421],
                        'cr_account_ids': [111, 112, 113],
                        'is_inverted_result': True,
                        'is_formula_active': False,
                        'account_ids': [],
                        'is_positive_difference': False,
                        'python_formular': ''
                    },
                ]
            },
        ]

        # create configuration data for account_balance_sheet report
        IndirCashFlowConfig = self.env['indirect.cash.flow.config']
        account_env = self.env['account.account']

        for line in LINES_CONFIG_DATA:
            try:
                # create parent_id
                indir_cashflow_config = IndirCashFlowConfig.create(
                    {
                        'item': line.get('item', ''),
                        'code': line.get('code', '')
                    }
                )
                indirect_cashflow_config_id = indir_cashflow_config and \
                    indir_cashflow_config.id or False
                # create child_ids
                for child_item_config in line.get('child_ids', []):
                    dr_account_ids = []
                    cr_account_ids = []
                    account_ids = []

                    for acc in child_item_config.get(
                            'dr_account_ids', []):
                        dr_account_ids += account_env.search(
                            [('code',
                              '=like',
                              '%s%%' % acc)]).ids

                    for acc in child_item_config.get(
                            'cr_account_ids', []):
                        cr_account_ids += account_env.search(
                            [('code',
                              '=like',
                              '%s%%' % acc)]).ids

                    for acc in child_item_config.get('account_ids', []):
                        account_ids += account_env.search(
                            [('code',
                              '=like', '%s%%' % acc)]).ids

                    dr_account_ids = [(6, 0, dr_account_ids)]
                    cr_account_ids = [(6, 0, cr_account_ids)]
                    account_ids = [(6, 0, account_ids)]

                    child_item_config.update(
                        {
                            'dr_account_ids': dr_account_ids,
                            'cr_account_ids': cr_account_ids,
                            'account_ids': account_ids,
                            'parent_id': indirect_cashflow_config_id
                        }
                    )

                    IndirCashFlowConfig.create(child_item_config)
            except:
                pass

        logging.info('END -- create indirect cashflow config data ...')
        return True
