from odoo import fields, models, api
from datetime import timedelta, datetime, date
from dateutil import relativedelta
from xlsxwriter.utility import xl_range
from xlsxwriter.utility import xl_rowcol_to_cell
import calendar
from odoo.tools.misc import DEFAULT_SERVER_DATE_FORMAT
import requests

class POBalanceReportXlsx(models.AbstractModel):
    _name = 'report.suntech_custom.po_balance_report_xls'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, wizard):
        orderObjSale = self.env['sale.order']

        for obj in wizard:
            start_date = datetime(year = obj.start_date.year, month = obj.start_date.month, day = obj.start_date.day, hour = 00, minute = 0, second = 1)
            end_date = datetime(year = obj.end_date.year, month = obj.end_date.month, day = obj.end_date.day, hour = 23, minute = 59, second = 59)

            if obj.is_st_plastics:
                main_search_sale = orderObjSale.search([('date_order','>=',start_date),('date_order','<=',end_date),('delivery_to_id','=',obj.shipping_id.id), ('state', '!=', 'cancel')], order='id asc')
                cell_format = workbook.add_format()

                title_parent = workbook.add_format({'bold':True, 'font_size':16})
                title = workbook.add_format({'bold':True, 'border':1, 'align': 'center', 'valign': 'vcenter', 'font_size':8})
                table = workbook.add_format({'border':1, 'font_size':9})
                table_center = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_percent = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_currency_idr = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_currency_usd = workbook.add_format({'align': 'center','border':1, 'font_size':9,'num_format': '$#,##0.00'})

                table_percent.set_num_format(10)
                table_currency_idr.set_num_format('"Rp" #,##0.00')

                sheet = workbook.add_worksheet('SO Balance Report')

                

                if obj.filter_by == 'all':
                    sheet.set_column('A:A', 5)
                    sheet.set_column('B:B', 10)
                    sheet.set_column('C:C', 8)
                    sheet.set_column('D:C', 12)
                    sheet.set_column('E:E', 10)
                    sheet.set_column('F:F', 10)
                    sheet.set_column('G:G', 30)
                    sheet.set_column('H:H', 18)
                    sheet.set_column('I:I', 25)
                    sheet.set_column('J:J', 8)
                    sheet.set_column('K:K', 12)
                    sheet.set_column('L:L', 12)
                    sheet.set_column('M:M', 10)
                    sheet.set_column('N:N', 10)
                    sheet.set_column('O:O', 12)
                    sheet.set_column('P:P', 12)
                    sheet.set_column('Q:Q', 12)
                    sheet.set_column('R:R', 12)
                    sheet.set_column('S:S', 18)

                    sheet.write('A1', 'No', title)
                    sheet.write('B1', 'PO Date', title)
                    sheet.write('C1', 'PO No', title)
                    sheet.write('D1', 'PO Reference', title)
                    sheet.write('E1', 'Required Date', title)
                    sheet.write('F1', 'Requestor', title)
                    sheet.write('G1', 'Supplier', title)
                    sheet.write('H1', 'Part Code', title)
                    sheet.write('I1', 'Part Name', title)
                    sheet.write('J1', 'Unit', title)
                    sheet.write('K1', 'Ordered', title)
                    sheet.write('L1', 'DO No', title)
                    sheet.write('M1', 'Do Date', title)
                    sheet.write('N1', 'DO Qty', title)
                    sheet.write('O1', 'OS', title)
                    sheet.write('P1', 'Status', title)
                    sheet.write('Q1', '%', title)
                    sheet.write('R1', 'Unit Price', title)
                    sheet.write('S1', 'Amount', title)

                    no = 1
                    row = 1
                    list_no = []
                    list_product = []

                    filter_all = self.set_all_data_master(main_search_sale, start_date, end_date)

                    for data in filter_all:
                        
                        percent = 0
                        if data['delivery'] and data['line'].product_uom_qty != 0:
                            if data['delivery'].product_qty != 0:
                                percent = (data['delivery'].product_qty/data['line'].product_uom_qty)
                        
                        if data['line'].product_uom_qty == 0:
                            percent = 1

                        if obj.format_type == 'sales':
                            if data['line'] not in list_no:
                                list_no.append(data['line'])
                                sheet.write(row, 0, data['no'], table_center)
                                sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 2, data['sale'].name, table_center)
                                sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, data['sale'].delivery_to_id.name, table_center)
                                sheet.write(row, 7, data['line'].product_id.default_code, table)
                                sheet.write(row, 8, data['line'].product_id.name, table)
                                sheet.write(row, 9, data['line'].product_uom.name, table_center)
                                sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                                sheet.write(row, 16, percent, table_percent)
                                if data['sale'].currency_id.name == 'IDR':
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                                else:
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)
                            else:
                                sheet.write(row, 0, '', table_center)
                                sheet.write(row, 1, '', table_center)
                                sheet.write(row, 2, '', table_center)
                                sheet.write(row, 3, '', table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, '', table_center)
                                sheet.write(row, 7, '', table)
                                sheet.write(row, 8, '', table_center)
                                sheet.write(row, 9, '', table_center)
                                sheet.write(row, 10, '', table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                                sheet.write(row, 16, percent, table_percent)
                                if data['sale'].currency_id.name == 'IDR':
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                                else:
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)

                            row += 1

                        else:
                            sheet.write(row, 0, data['no'], table_center)
                            sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                            sheet.write(row, 2, data['sale'].name, table_center)
                            sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                            sheet.write(row, 4, '', table_center)
                            sheet.write(row, 5, '', table_center)
                            sheet.write(row, 6, data['sale'].delivery_to_id.name, table_center)
                            sheet.write(row, 7, data['line'].product_id.default_code, table)
                            sheet.write(row, 8, data['line'].product_id.name, table)
                            sheet.write(row, 9, data['line'].product_uom.name, table_center)
                            sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                            if not data['delivery']:
                                sheet.write(row, 11, '', table_center)
                                sheet.write(row, 12, '', table_center)
                                sheet.write(row, 13, '', table_center)
                            else:
                                sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 13, data['delivery'].product_qty, table_center)
                            sheet.write(row, 14, data['os'], table_center)
                            if data['os'] > 0:
                                sheet.write(row, 15, 'Running', table_center)
                            else:
                                sheet.write(row, 15, 'Done', table_center)
                            sheet.write(row, 16, percent, table_percent)
                            if data['sale'].currency_id.name == 'IDR':
                                sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                            else:
                                sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)

                            row += 1

                    
                else:
                    sheet.set_column('A:A', 5)
                    sheet.set_column('B:B', 10)
                    sheet.set_column('C:C', 8)
                    sheet.set_column('D:C', 12)
                    sheet.set_column('E:E', 10)
                    sheet.set_column('F:F', 10)
                    sheet.set_column('G:G', 30)
                    sheet.set_column('H:H', 18)
                    sheet.set_column('I:I', 25)
                    sheet.set_column('J:J', 8)
                    sheet.set_column('K:K', 12)
                    sheet.set_column('L:L', 12)
                    sheet.set_column('M:M', 10)
                    sheet.set_column('N:N', 10)
                    sheet.set_column('O:O', 12)
                    sheet.set_column('P:P', 12)

                    sheet.write('A1', 'No', title)
                    sheet.write('B1', 'PO Date', title)
                    sheet.write('C1', 'PO No', title)
                    sheet.write('D1', 'PO Reference', title)
                    sheet.write('E1', 'Required Date', title)
                    sheet.write('F1', 'Requestor', title)
                    sheet.write('G1', 'Supplier', title)
                    sheet.write('H1', 'Part Code', title)
                    sheet.write('I1', 'Part Name', title)
                    sheet.write('J1', 'Unit', title)
                    sheet.write('K1', 'Ordered', title)
                    sheet.write('L1', 'DO No', title)
                    sheet.write('M1', 'Do Date', title)
                    sheet.write('N1', 'DO Qty', title)
                    sheet.write('O1', 'OS', title)
                    sheet.write('P1', 'Status', title)

                    row = 1

                    filter_all = self.set_outstanding_data_master(main_search_sale, start_date, end_date)
                    line_list = []
                    
                    for data in filter_all:
                        if obj.format_type == 'sales':
                            if data['line'] not in line_list:
                                line_list.append(data['line'])
                                sheet.write(row, 0, data['no'], table_center)
                                sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 2, data['sale'].name, table_center)
                                sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, data['sale'].delivery_to_id.name, table_center)
                                sheet.write(row, 7, data['line'].product_id.default_code, table)
                                sheet.write(row, 8, data['line'].product_id.name, table)
                                sheet.write(row, 9, data['line'].product_uom.name, table_center)
                                sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                            else:
                                sheet.write(row, 0, '', table_center)
                                sheet.write(row, 1, '', table_center)
                                sheet.write(row, 2, '', table_center)
                                sheet.write(row, 3, '', table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, '', table_center)
                                sheet.write(row, 7, '', table)
                                sheet.write(row, 8, '', table)
                                sheet.write(row, 9, '', table_center)
                                sheet.write(row, 10, '', table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                            
                            row += 1
                        else:
                            sheet.write(row, 0, data['no'], table_center)
                            sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                            sheet.write(row, 2, data['sale'].name, table_center)
                            sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                            sheet.write(row, 4, '', table_center)
                            sheet.write(row, 5, '', table_center)
                            sheet.write(row, 6, data['sale'].delivery_to_id.name, table_center)
                            sheet.write(row, 7, data['line'].product_id.default_code, table)
                            sheet.write(row, 8, data['line'].product_id.name, table)
                            sheet.write(row, 9, data['line'].product_uom.name, table_center)
                            sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                            if not data['delivery']:
                                sheet.write(row, 11, '', table_center)
                                sheet.write(row, 12, '', table_center)
                                sheet.write(row, 13, '', table_center)
                            else:
                                sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 13, data['delivery'].product_qty, table_center)
                            sheet.write(row, 13, data['os'], table_center)
                            if data['os'] > 0:
                                sheet.write(row, 15, 'Running', table_center)
                            else:
                                sheet.write(row, 15, 'Done', table_center)
                            
                            row += 1
            else:
                main_search_sale = orderObjSale.search([('date_order','>=',start_date),('date_order','<=',end_date),('partner_id','=',obj.partner_id.id), ('state', '!=', 'cancel')], order='id asc')
                cell_format = workbook.add_format()

                title_parent = workbook.add_format({'bold':True, 'font_size':16})
                title = workbook.add_format({'bold':True, 'border':1, 'align': 'center', 'valign': 'vcenter', 'font_size':8})
                table = workbook.add_format({'border':1, 'font_size':9})
                table_center = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_percent = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_currency_idr = workbook.add_format({'align': 'center','border':1, 'font_size':9})
                table_currency_usd = workbook.add_format({'align': 'center','border':1, 'font_size':9,'num_format': '$#,##0.00'})

                table_percent.set_num_format(10)
                table_currency_idr.set_num_format('"Rp" #,##0.00')

                sheet = workbook.add_worksheet('SO Balance Report')

                

                if obj.filter_by == 'all':
                    sheet.set_column('A:A', 5)
                    sheet.set_column('B:B', 10)
                    sheet.set_column('C:C', 8)
                    sheet.set_column('D:C', 12)
                    sheet.set_column('E:E', 10)
                    sheet.set_column('F:F', 10)
                    sheet.set_column('G:G', 30)
                    sheet.set_column('H:H', 18)
                    sheet.set_column('I:I', 25)
                    sheet.set_column('J:J', 8)
                    sheet.set_column('K:K', 12)
                    sheet.set_column('L:L', 12)
                    sheet.set_column('M:M', 10)
                    sheet.set_column('N:N', 10)
                    sheet.set_column('O:O', 12)
                    sheet.set_column('P:P', 12)
                    sheet.set_column('Q:Q', 12)
                    sheet.set_column('R:R', 12)
                    sheet.set_column('S:S', 18)

                    sheet.write('A1', 'No', title)
                    sheet.write('B1', 'PO Date', title)
                    sheet.write('C1', 'PO No', title)
                    sheet.write('D1', 'PO Reference', title)
                    sheet.write('E1', 'Required Date', title)
                    sheet.write('F1', 'Requestor', title)
                    sheet.write('G1', 'Supplier', title)
                    sheet.write('H1', 'Part Code', title)
                    sheet.write('I1', 'Part Name', title)
                    sheet.write('J1', 'Unit', title)
                    sheet.write('K1', 'Ordered', title)
                    sheet.write('L1', 'DO No', title)
                    sheet.write('M1', 'Do Date', title)
                    sheet.write('N1', 'DO Qty', title)
                    sheet.write('O1', 'OS', title)
                    sheet.write('P1', 'Status', title)
                    sheet.write('Q1', '%', title)
                    sheet.write('R1', 'Unit Price', title)
                    sheet.write('S1', 'Amount', title)

                    no = 1
                    row = 1
                    list_no = []
                    list_product = []
                    print(111111111111)
                    filter_all = self.set_all_data_master(main_search_sale, start_date, end_date)
                    print(filter_all,2222222222)

                    for data in filter_all:
                        print(333333333333333)
                        percent = 0
                        if data['delivery'] and data['line'].product_uom_qty != 0:
                            if data['delivery'].product_qty != 0:
                                percent = (data['delivery'].product_qty/data['line'].product_uom_qty)
                        
                        if data['line'].product_uom_qty == 0:
                            percent = 1

                        if obj.format_type == 'sales':
                            if data['line'] not in list_no:
                                list_no.append(data['line'])
                                sheet.write(row, 0, data['no'], table_center)
                                sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 2, data['sale'].name, table_center)
                                sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, data['sale'].partner_id.name, table_center)
                                sheet.write(row, 7, data['line'].product_id.default_code, table)
                                sheet.write(row, 8, data['line'].product_id.name, table)
                                sheet.write(row, 9, data['line'].product_uom.name, table_center)
                                sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                                sheet.write(row, 16, percent, table_percent)
                                if data['sale'].currency_id.name == 'IDR':
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                                else:
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)
                            else:
                                sheet.write(row, 0, '', table_center)
                                sheet.write(row, 1, '', table_center)
                                sheet.write(row, 2, '', table_center)
                                sheet.write(row, 3, '', table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, '', table_center)
                                sheet.write(row, 7, '', table)
                                sheet.write(row, 8, '', table_center)
                                sheet.write(row, 9, '', table_center)
                                sheet.write(row, 10, '', table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                                sheet.write(row, 16, percent, table_percent)
                                if data['sale'].currency_id.name == 'IDR':
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                                else:
                                    sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                    sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)

                            row += 1

                        else:
                            sheet.write(row, 0, data['no'], table_center)
                            sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                            sheet.write(row, 2, data['sale'].name, table_center)
                            sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                            sheet.write(row, 4, '', table_center)
                            sheet.write(row, 5, '', table_center)
                            sheet.write(row, 6, data['sale'].partner_id.name, table_center)
                            sheet.write(row, 7, data['line'].product_id.default_code, table)
                            sheet.write(row, 8, data['line'].product_id.name, table)
                            sheet.write(row, 9, data['line'].product_uom.name, table_center)
                            sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                            if not data['delivery']:
                                sheet.write(row, 11, '', table_center)
                                sheet.write(row, 12, '', table_center)
                                sheet.write(row, 13, '', table_center)
                            else:
                                sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 13, data['delivery'].product_qty, table_center)
                            sheet.write(row, 14, data['os'], table_center)
                            if data['os'] > 0:
                                sheet.write(row, 15, 'Running', table_center)
                            else:
                                sheet.write(row, 15, 'Done', table_center)
                            sheet.write(row, 16, percent, table_percent)
                            if data['sale'].currency_id.name == 'IDR':
                                sheet.write(row, 17, data['line'].price_unit, table_currency_idr)
                                sheet.write(row, 18, data['line'].price_subtotal, table_currency_idr)
                            else:
                                sheet.write(row, 17, data['line'].price_unit, table_currency_usd)
                                sheet.write(row, 18, data['line'].price_subtotal, table_currency_usd)

                            row += 1

                    
                else:
                    sheet.set_column('A:A', 5)
                    sheet.set_column('B:B', 10)
                    sheet.set_column('C:C', 8)
                    sheet.set_column('D:C', 12)
                    sheet.set_column('E:E', 10)
                    sheet.set_column('F:F', 10)
                    sheet.set_column('G:G', 30)
                    sheet.set_column('H:H', 18)
                    sheet.set_column('I:I', 25)
                    sheet.set_column('J:J', 8)
                    sheet.set_column('K:K', 12)
                    sheet.set_column('L:L', 12)
                    sheet.set_column('M:M', 10)
                    sheet.set_column('N:N', 10)
                    sheet.set_column('O:O', 12)
                    sheet.set_column('P:P', 12)

                    sheet.write('A1', 'No', title)
                    sheet.write('B1', 'PO Date', title)
                    sheet.write('C1', 'PO No', title)
                    sheet.write('D1', 'PO Reference', title)
                    sheet.write('E1', 'Required Date', title)
                    sheet.write('F1', 'Requestor', title)
                    sheet.write('G1', 'Supplier', title)
                    sheet.write('H1', 'Part Code', title)
                    sheet.write('I1', 'Part Name', title)
                    sheet.write('J1', 'Unit', title)
                    sheet.write('K1', 'Ordered', title)
                    sheet.write('L1', 'DO No', title)
                    sheet.write('M1', 'Do Date', title)
                    sheet.write('N1', 'DO Qty', title)
                    sheet.write('O1', 'OS', title)
                    sheet.write('P1', 'Status', title)

                    row = 1

                    filter_all = self.set_outstanding_data_master(main_search_sale, start_date, end_date)
                    line_list = []
                    
                    for data in filter_all:
                        if obj.format_type == 'sales':
                            if data['line'] not in line_list:
                                line_list.append(data['line'])
                                sheet.write(row, 0, data['no'], table_center)
                                sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 2, data['sale'].name, table_center)
                                sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, data['sale'].partner_id.name, table_center)
                                sheet.write(row, 7, data['line'].product_id.default_code, table)
                                sheet.write(row, 8, data['line'].product_id.name, table)
                                sheet.write(row, 9, data['line'].product_uom.name, table_center)
                                sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                            else:
                                sheet.write(row, 0, '', table_center)
                                sheet.write(row, 1, '', table_center)
                                sheet.write(row, 2, '', table_center)
                                sheet.write(row, 3, '', table_center)
                                sheet.write(row, 4, '', table_center)
                                sheet.write(row, 5, '', table_center)
                                sheet.write(row, 6, '', table_center)
                                sheet.write(row, 7, '', table)
                                sheet.write(row, 8, '', table)
                                sheet.write(row, 9, '', table_center)
                                sheet.write(row, 10, '', table_center)
                                if not data['delivery']:
                                    sheet.write(row, 11, '', table_center)
                                    sheet.write(row, 12, '', table_center)
                                    sheet.write(row, 13, '', table_center)
                                else:
                                    sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                    sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                    sheet.write(row, 13, data['delivery'].product_qty, table_center)
                                sheet.write(row, 14, data['os'], table_center)
                                if data['os'] > 0:
                                    sheet.write(row, 15, 'Running', table_center)
                                else:
                                    sheet.write(row, 15, 'Done', table_center)
                            
                            row += 1
                        else:
                            sheet.write(row, 0, data['no'], table_center)
                            sheet.write(row, 1, data['sale'].date_order.strftime('%d/%m/%Y'), table_center)
                            sheet.write(row, 2, data['sale'].name, table_center)
                            sheet.write(row, 3, data['sale'].client_order_ref, table_center)
                            sheet.write(row, 4, '', table_center)
                            sheet.write(row, 5, '', table_center)
                            sheet.write(row, 6, data['sale'].partner_id.name, table_center)
                            sheet.write(row, 7, data['line'].product_id.default_code, table)
                            sheet.write(row, 8, data['line'].product_id.name, table)
                            sheet.write(row, 9, data['line'].product_uom.name, table_center)
                            sheet.write(row, 10, data['line'].product_uom_qty, table_center)
                            if not data['delivery']:
                                sheet.write(row, 11, '', table_center)
                                sheet.write(row, 12, '', table_center)
                                sheet.write(row, 13, '', table_center)
                            else:
                                sheet.write(row, 11, data['delivery'].picking_id.name, table_center)
                                sheet.write(row, 12, data['delivery'].picking_id.scheduled_date.strftime('%d/%m/%Y'), table_center)
                                sheet.write(row, 13, data['delivery'].product_qty, table_center)
                            sheet.write(row, 13, data['os'], table_center)
                            if data['os'] > 0:
                                sheet.write(row, 15, 'Running', table_center)
                            else:
                                sheet.write(row, 15, 'Done', table_center)
                            
                            row += 1
                
                

    def get_delivery_data(self, line, start_date, end_date):
        query = '''
            SELECT 
                sm.id
            FROM stock_move sm 
            LEFT JOIN stock_picking sp 
            ON 
                sm.picking_id = sp.id 
            WHERE 
                sp.state != 'cancel' AND
                sp.scheduled_date BETWEEN '%s' AND '%s' AND 
                sm.sol_id = %s AND
                sp.is_pick_so = True
        '''%(start_date, end_date, line.id)
        
        self.env.cr.execute(query)
        delivery_list = self.env.cr.fetchall()
        return delivery_list
    
    def check_available_delivery(self, line, start_date, end_date):
        delivery_value = []
        delivery_list = self.get_delivery_data(line, start_date, end_date)
        move_obj = self.env['stock.move']
        
        if delivery_list:
            for delivery in delivery_list:
                delivery_id = move_obj.browse(delivery[0])
                delivery_value.append(delivery_id)
        
        return delivery_value
    
    def check_outstanding_balance(self, line, start_date, end_date):
        is_outstanding = False
        total = line.product_uom_qty
        delivery_value = []
        delivery_list = self.get_delivery_data(line, start_date, end_date)
        move_obj = self.env['stock.move']
        if delivery_list:
            for delivery in delivery_list:
                delivery_id = move_obj.browse(delivery[0])
                total -= delivery_id.product_qty

        if total > 0:
            is_outstanding = True
        
        return is_outstanding


    def set_all_data_master(self, main_data, start_date, end_date):
        data = []
        no = 1
        for sale in main_data:
            for line in sale.order_line:
                available_delivery = self.check_available_delivery(line, start_date, end_date)
                os = line.product_uom_qty
                if len(available_delivery) > 0:
                    for delivery in available_delivery:
                        os -= delivery.product_qty
                        data.append({
                            'no':no,
                            'sale':sale,
                            'line':line,
                            'delivery':delivery,
                            'os':os
                        })
                else:
                    data.append({
                        'no':no,
                        'sale':sale,
                        'line':line,
                        'delivery':False,
                        'os':os
                    })
            no += 1
        return data
    
    def set_outstanding_data_master(self, main_search_sale, start_date, end_date):
        data = []
        no = 1
        for sale in main_search_sale:
            for line in sale.order_line:
                is_outstanding = self.check_outstanding_balance(line, start_date, end_date)
                if is_outstanding:
                    available_delivery = self.check_available_delivery(line, start_date, end_date)
                    os = line.product_uom_qty
                    if available_delivery:
                        for delivery in available_delivery:
                            os -= delivery.product_qty
                            data.append({
                                'no':no,
                                'sale':sale,
                                'line':line,
                                'delivery':delivery,
                                'os':os
                            })
                    else:
                        data.append({
                        'no':no,
                        'sale':sale,
                        'line':line,
                        'delivery':False,
                        'os':os
                    })
                
            no += 1
        return data
