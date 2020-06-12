from amzn_parser_utils import dkey_to_float, is_windows_machine, simplify_date, col_to_letter
from amzn_parser_constants import TEMPLATE_SHEET_MAPPING
import logging
import openpyxl
import os
# Delete later
import json
import subprocess

# GLOBAL VARIABLES
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
SUMMARY_SHEET_NAME = 'Summary'
SHEET_HEADERS = list(TEMPLATE_SHEET_MAPPING.keys())


class OrdersReport():
    '''accepts export data dictionary and output file path as arguments, creates individual
    sheets for region & currency based order segregation, creates formatted xlsx report file.
    Error handling is present outside of this class.
    
    Main method: export() - creates individual sheets, pushes selected data from corresponding orders;
    creates summary sheet, calculates regional / currency based totals'''
    
    def __init__(self, export_obj : dict):
        self.export_obj = export_obj
        # {'EU EUR': {'Total': 2066, 'Count':69, 'Average': 30.61}}
        self.totals = {}
    
    def unpack_export_obj(self):
        '''generator yielding self.export_obj dict KEY1 (region) and nested KEY2 (currency) at a time'''
        for region in self.export_obj:
            for currency in self.export_obj[region]:
                yield region, currency    

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
            elif d_key in ['purchase-date', 'payments-date']:
                simpler_date = simplify_date(order_dict[d_key])
                d_with_output_keys[header] = simpler_date
            else:
                d_with_output_keys[header] = order_dict[TEMPLATE_SHEET_MAPPING[header]]
        return d_with_output_keys

    def _data_to_sheet(self, ws_name : str, orders_data: list):
        '''creates new ws_name sheet and fills it with orders_data argument data'''
        self._create_sheet(ws_name)
        active_ws = self.wb[ws_name]
        active_ws.freeze_panes = active_ws['A2']
        self._fill_sheet(active_ws, orders_data)
        self._adjust_col_widths(active_ws, self.col_widths)

    def _create_sheet(self, ws_name : str):
        '''creates new sheet obj with name ws_name'''
        self.wb.create_sheet(title=ws_name)
        self.col_widths = {}

    def _fill_sheet(self, active_ws, orders_data:list):
        '''writes headers, and corresponding orders data to active_ws, resizes columns'''
        self._write_headers(active_ws, SHEET_HEADERS)
        self._write_sheet_orders(active_ws, SHEET_HEADERS, orders_data)

    def _write_headers(self, worksheet, headers):
        for col, header in enumerate(headers):
            self._update_col_widths(col, header)
            worksheet.cell(1, col + 1).value = header    

    @staticmethod
    def range_generator(orders_data, headers):
        for row, _ in enumerate(orders_data):
            for col, _ in enumerate(headers):
                yield row, col
    
    def _write_sheet_orders(self, worksheet, headers, orders_data):
        for row, col in self.range_generator(orders_data, headers):
            working_dict = orders_data[row]
            key_pointer = headers[col]
            self._update_col_widths(col, str(working_dict[key_pointer]))
            # offsets due to excel vs python numbering  + headers in row 1
            worksheet.cell(row + 2, col + 1).value = working_dict[key_pointer]

    def _update_col_widths(self, col : int, cell_value : str):
        '''runs on each cell. Forms a dictionary {'A':30, 'B':15...} for max column widths in worksheet (width as length of max cell)'''
        col_letter = col_to_letter(col)
        if col_letter in self.col_widths:
            # check for length, update if current cell length exceeds current entry for column
            if len(cell_value) > self.col_widths[col_letter]:
                self.col_widths[col_letter] = len(cell_value)
        else:
            self.col_widths[col_letter] = len(cell_value)

    def _adjust_col_widths(self, worksheet, col_widths : dict):
        '''iterates over {'A':30, 'B':40, 'C':35...} dict to resize worksheets' column widths'''
        for col_letter in col_widths:
            adjusted_width = ((col_widths[col_letter]) + 2 * 1.05 )
            worksheet.column_dimensions[col_letter].width = adjusted_width
    
    def fill_summary(self):
        '''Forms a summary statistics sheet / report'''
        active_ws = self.wb[SUMMARY_SHEET_NAME]
        self.write_totals_table(active_ws)
        self.write_daily_breakdown_table(active_ws)        
        
    def write_totals_table(self, ws):
        '''fills out first table in main sheet of report.
        Table height is fixed, but width is flexible depending on data segments'''
        # Write vertical headers (Total, Orders, Average in A3:A5)
        any_values_dict = list(self.totals.values())[0]
        for row, values_header in enumerate(any_values_dict.keys()):
            ws.cell(row + 3, 1).value = values_header
        # Fill segment name and values (range B2:[n]5):
        for col, (segment, values) in enumerate(self.totals.items()):
            ws.cell(2, col + 2).value = segment
            ws.cell(3, col + 2).value = values['Total']
            ws.cell(4, col + 2).value = values['Orders']
            ws.cell(5, col + 2).value = values['Average']
    
    def write_daily_breakdown_table(self, ws):
        '''fills out second table, which breaks down daily statistics for each segment'''
        ws.cell(10, 3).value = 'Somewhere here will come the more powerful table.'
        # pass

    def get_segment_total_count_avg(self, segment_name : str, segment_orders : list) -> dict:
        '''adds to self.totals a new key (segment name), calculated total, count and average'''
        s_total = self.get_total(segment_orders)
        s_count = len(segment_orders)
        s_avg = round(s_total / s_count, 2)
        return {'Total' : s_total, 'Orders' : s_count, 'Average' : s_avg}

    @staticmethod
    def get_total(orders : list) -> float:
        '''Returns a sum of all orders' item-price + item-tax in list of order dicts'''
        total = 0
        for order in orders:
            total += dkey_to_float(order, 'item-price') + dkey_to_float(order, 'item-tax')
        return total

    def export(self, wb_name : str):
        '''unpacks data object and pushes orders data to appropriate sheets, forms a summary sheet, saves new workbook'''
        try:
            self.wb = openpyxl.Workbook()
            ws = self.wb.active
            ws.title = SUMMARY_SHEET_NAME
            # Unpacking data to sheets:
            for region, currency in self.unpack_export_obj():
                segment_orders = self.export_obj[region][currency]
                self.totals[f'{region} {currency}'] = self.get_segment_total_count_avg(f'{region} {currency}', segment_orders)
                sheet_ready_data = self.apply_orders_to_sheet_format(segment_orders)
                self._data_to_sheet(f'{region} {currency}', sheet_ready_data)
            self.fill_summary()
        except Exception as e:
            print(f'Unknown error while creating excel report {os.path.basename(wb_name)}. Err: {e}')
        finally:
            self.wb.save(wb_name)
            self.wb.close()


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
    report.export(excel_file)
    print('File saved. Check.')
    open_file(excel_file)


if __name__ == "__main__":
    subprocess.call(['rm', 'EXCEL REPORT.xlsx'])
    run()
    # pass