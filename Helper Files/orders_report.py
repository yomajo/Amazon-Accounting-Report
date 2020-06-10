from amzn_parser_utils import dkey_to_float ,is_windows_machine
from amzn_parser_constants import TEMPLATE_SHEET_MAPPING
import logging
import openpyxl
# Delete later
import json
import subprocess

# GLOBAL VARIABLES
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
SHEET_HEADERS = list(TEMPLATE_SHEET_MAPPING.keys())


class OrdersReport():
    '''accepts export data dictionary and output file path as arguments, creates individual
    sheets for region & currency based order segregation, creates formatted xlsx report file.
    Error handling is present outside of this class.
    
    Main method: export() - creates individual sheets, pushes selected data from corresponding orders;
    creates summary sheet, calculates regional / currency'''
    
    def __init__(self, export_obj : dict):
        self.export_obj = export_obj
    
    def unpack_export_obj(self):
        '''generator yielding self.export_obj dict KEY1 (region) and nested KEY2 (currency) at a time'''
        for region in self.export_obj:
            for currency in self.export_obj[region]:
                # print(f'Region {region} contains currency: {currency}')
                # print(f'Inside this contains object of type: {type(self.export_obj[region][currency])}, which has {len(self.export_obj[region][currency])} members')
                yield region, currency

    def export_wb(self, wb_name : str):
        '''creates a workbook from thin air'''
        try:
            self.wb = openpyxl.Workbook()
            ws = self.wb.active
            ws.title = 'Summary'
            # Unpacking data to sheets:
            for region, currency in self.unpack_export_obj():
                working_orders = self.export_obj[region][currency]
                sheet_ready_data = self.apply_orders_to_sheet_format(working_orders)
                self._data_to_sheet(f'{region} {currency}', sheet_ready_data)

        except Exception as e:
            print(f'Errors while creating excel file. Err: {e}')
        finally:
            self.wb.save(wb_name)
            self.wb.close()
    

    def _data_to_sheet(self, ws_name : str, orders_data: list):
        '''creates new ws_name sheet and fills it with orders_data argument'''
        self._create_sheet(ws_name)
        active_ws = self.wb[ws_name]
        active_ws.freeze_panes = active_ws['A2']
        self._fill_sheet(active_ws, orders_data)


    def _create_sheet(self, ws_name : str):
        '''creates new sheet obj with name ws_name'''
        self.wb.create_sheet(title=ws_name)

    def _fill_sheet(self, active_ws, orders_data:list):
        '''writes headers, and corresponding orders data to active_ws'''
        self._write_headers(active_ws, SHEET_HEADERS)
        self._write_sheet_orders(active_ws, SHEET_HEADERS, orders_data)

# --------------------------------------------------------------------------------

    def apply_orders_to_sheet_format(self, raw_orders_data : list):
        '''reduces input data to that needed in output data sheet'''
        ready_orders_data = []
        for order_dict in raw_orders_data:
            reduced_order_dict = self.get_mapped_dict(order_dict)
            ready_orders_data.append(reduced_order_dict)
        return ready_orders_data

    def get_mapped_dict(self, order_dict : dict):
        d_with_output_keys = {}
        for header, d_key in TEMPLATE_SHEET_MAPPING.items():
            if d_key in ['item-price', 'item-tax', 'shipping-price', 'shipping-tax']:
                d_value_as_float = dkey_to_float(order_dict, d_key)
                d_with_output_keys[header] = d_value_as_float
            else:
                d_with_output_keys[header] = order_dict[TEMPLATE_SHEET_MAPPING[header]]
        return d_with_output_keys
    

    #  ---------------------------------------------------------------------------

    @staticmethod
    def _write_headers(worksheet, headers):
        for col, header in enumerate(headers, 1):
            worksheet.cell(1, col).value = header    

    @staticmethod
    def range_generator(orders_data, headers):
        for row, _ in enumerate(orders_data):
            for col, _ in enumerate(headers):
                yield row, col
    
    def _write_sheet_orders(self, worksheet, headers, orders_data):
        for row, col in self.range_generator(orders_data, headers):
            working_dict = orders_data[row]
            # Line saving to functions:
            # fake_pointer = TEMPLATE_SHEET_MAPPING[headers[col]]
            # print(f'I would get value with this key in orders_dict: {fake_pointer}')

            key_pointer = headers[col]
            # offsets due to excel vs python numbering  + headers in row 1
            worksheet.cell(row + 2, col + 1).value = working_dict[key_pointer]


    def export(self, report_path : str):
        # Silencing error ------------
        self.etonas_orders = []
        # Silencing error ------------
        reheaded_etonas_orders = self.apply_orders_to_sheet_format(self.etonas_orders)
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            self._write_headers(ws, SHEET_HEADERS)
            self._write_sheet_orders(ws, SHEET_HEADERS, reheaded_etonas_orders)
            # self.adjust_col_widths(ws)
        except Exception as e:
            print(f'Error occured while creating report. Err: {e}. Saving, exiting')
        finally:
            wb.save(report_path)

    # def adjust_col_widths(self, worksheet):
    #     for col in worksheet.columns:
    #         max_length = 0
    #         col_letter = col[0].column_letter
    #         for cell in col:
    #             try:
    #                 if len(str(cell.value)) > max_length:
    #                     max_length = len(cell.value)
    #             except:
    #                 pass
    #         adjusted_width = (max_length + 2) * 1.1
    #         worksheet.column_dimensions[col_letter].width = adjusted_width

    def fill_summary(self):
        pass

def read_json(fname:str):
    with open(fname, 'r', encoding='utf-8') as f:
        contents = json.load(f)
    return contents

def open_file(fname:str):
    subprocess.call(['xdg-open',fname])

def run():
    excel_file = 'EXCEL REPORT.xlsx'
    data_obj = read_json('export_data.json')
    report = OrdersReport(data_obj)
    report.export_wb(excel_file)
    print('File saved. Or not. Check.')

    open_file(excel_file)
    # report.export('EXCEL REPORT.xlsx')


if __name__ == "__main__":
    subprocess.call(['rm', 'EXCEL REPORT.xlsx'])
    run()
    # pass