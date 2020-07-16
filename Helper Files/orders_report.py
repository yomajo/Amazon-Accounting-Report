from amzn_parser_constants import TEMPLATE_SHEET_MAPPING, SUMMARY_HEADERS
from amzn_parser_utils import simplify_date, col_to_letter
from collections import defaultdict
import openpyxl
import os

# GLOBAL VARIABLES
SUMMARY_SHEET_NAME = 'Summary'
BOLD_STYLE = openpyxl.styles.Font(bold=True, name='Calibri')
BACKGROUND_COLOR_STYLE = openpyxl.styles.PatternFill(fgColor='D6DEFF',fill_type='solid')
THIN_BORDER = openpyxl.styles.Side(border_style='thin')
SHEET_HEADERS = list(TEMPLATE_SHEET_MAPPING.keys())
REPORT_START_ROW = 1
REPORT_START_COL = 1


class OrdersReport():
    '''accepts export data dictionary and output file path as arguments, creates individual
    sheets for region & currency based order segregation, creates formatted xlsx report file.
    Error handling is present outside of this class.

    Expected input obj format: 
        {eu_orders: {currency1: [order1, order2, order...], currency2: [order1, order2, order...], ...},
        non_eu_orders: {currency1: [order1, order2, order...], currency2: [order1, order2, order...] ...},
        ...}
    
    Main method: export() - creates individual sheets, pushes selected data from corresponding orders;
    creates summary sheet, calculates regional / currency based totals'''
    
    def __init__(self, export_obj : dict):
        self.export_obj = self._clean_incoming_data(export_obj)
        self.__get_report_objs()

    def _clean_incoming_data(self, export_obj : dict) -> dict:
        '''returns cleaned data in same structure as class input obj'''
        for region, currency in self._unpack_export_obj(export_obj):
            export_obj[region][currency] = self._correct_orders_numbers_dates(export_obj[region][currency])
        return export_obj

    @staticmethod
    def _unpack_export_obj(export_obj : dict):
        '''generator yielding self.export_obj dict KEY1 (region) and nested KEY2 (currency) at a time'''
        for region in export_obj:
            for currency in export_obj[region]:
                yield region, currency    
    
    @staticmethod
    def _correct_orders_numbers_dates(orders : list) -> list:
        '''date and numbers data cleaning for report:
        - original date format (2020-04-16T10:07:16+00:00) simplified to YYYY-MM-DD
        - all numbers of interest converted to floats'''
        for order in orders:
            # Simplifying order dates:
            order['purchase-date'] = simplify_date(order['purchase-date'])
            order['payments-date'] = simplify_date(order['payments-date'])
            # Converting numbers stored as strings to floats:
            order['item-price'] = float(order['item-price'])
            order['item-tax'] = float(order['item-tax'])
            order['shipping-price'] = float(order['shipping-price'])
            order['shipping-tax'] = float(order['shipping-tax'])
        return orders
    
    def __get_report_objs(self):
        '''prepares cls variables: self.segments_orders_obj, self.summary_table_obj for excel report workbook filling'''
        self.segments_orders_obj = self._get_segments_orders_obj(self.export_obj)
        self.summary_table_obj = self._get_summary_table_obj(self.export_obj)

    def _get_segments_orders_obj(self, export_obj : dict) -> dict:
        '''returns dict of dicts for each segment (sheets) and corresponding list of orders
        returned object: {'EU EUR' : [order1, order2, ...], 'NON-EU GBP' : [order11, order 22...], ...}'''
        segments_orders = {}
        for region, currency in self._unpack_export_obj(export_obj):
            segments_orders[f'{region} {currency}'] = export_obj[region][currency]
        return segments_orders

    def _get_summary_table_obj(self, export_obj : dict) -> dict:
        '''Returns nestd object based on 1. currency 2. order dates. Example:
        {'EUR':{'date1':[order1, order2...], 'date2':[order3, order4...], ...},
        'GBP':{'date1':[order1, order2...], 'date2':[order1, order2...], ...}, ...}'''
        summary_table_obj = defaultdict(dict)
        currency_based = defaultdict(list)
        # 1st loop forms currency (key) based obj. Output: {'currency1':[order1, order2, ...], 'currency2':[order2, order3...],...}
        for region, currency in self._unpack_export_obj(export_obj):
            currency_based[currency] = currency_based[currency] + export_obj[region][currency]
        # 2nd loop - forms {'currency':{'date1':[order1, order2, ...], 'date2':[order1, order2, ...], ...}, 'currency2':{...}}
        for currency in currency_based:
            summary_table_obj[currency] = self._get_payment_date_obj(currency_based[currency])
        return summary_table_obj

    @staticmethod
    def _get_payment_date_obj(orders : list) -> dict:
        '''forms a dictionary based on dates in list of order dicts. Returns payment dates as keys and orders list as values.
        Output format: {{'YYYY-MM-D1':[order1, order2, ...]},
                        {'YYYY-MM-D2':[order1, order2, ...]}, ...}'''
        payment_date_orders = defaultdict(list)
        for order in orders:
            payment_date_orders[order['payments-date']].append(order)
        return payment_date_orders

    def _data_to_sheet(self, ws_name : str, orders_data: list):
        '''creates new ws_name sheet and fills it with orders_data argument data'''
        self.__create_sheet(ws_name)
        active_ws = self.wb[ws_name]
        active_ws.freeze_panes = active_ws['A2']
        self._fill_sheet(active_ws, orders_data)
        self._adjust_col_widths(active_ws, self.col_widths)

    def __create_sheet(self, ws_name : str):
        '''creates new sheet obj with name ws_name'''
        self.wb.create_sheet(title=ws_name)
        self.col_widths = {}

    def _fill_sheet(self, active_ws, orders_data:list):
        '''writes headers, and corresponding orders data to active_ws, resizes columns'''
        self._write_headers(active_ws)
        self._write_sheet_orders(active_ws, orders_data)

    def _write_headers(self, ws):
        for col, header in enumerate(SHEET_HEADERS):
            self.__update_col_widths(col, header)
            ws.cell(1, col + 1).value = header

    def __update_col_widths(self, col : int, cell_value : str, zero_indexed=True):
        '''runs on each cell. Forms a dictionary {'A':30, 'B':15...} for max column widths in worksheet (width as length of max cell)'''
        col_letter = col_to_letter(col, zero_indexed=zero_indexed)
        if col_letter in self.col_widths:
            # check for length, update if current cell length exceeds current entry for column
            if len(cell_value) > self.col_widths[col_letter]:
                self.col_widths[col_letter] = len(cell_value)
        else:
            self.col_widths[col_letter] = len(cell_value)

    def _write_sheet_orders(self, ws, orders_data : list):
        for row, col in self.range_generator(orders_data, SHEET_HEADERS):
            working_dict = orders_data[row]
            key_pointer = TEMPLATE_SHEET_MAPPING[SHEET_HEADERS[col]]
            self.__update_col_widths(col, str(working_dict[key_pointer]))
            # offsets due to excel vs python numbering  + headers in row 1
            ws.cell(row + 2, col + 1).value = working_dict[key_pointer]

    @staticmethod
    def range_generator(orders_data, headers):
        for row, _ in enumerate(orders_data):
            for col, _ in enumerate(headers):
                yield row, col

    def _adjust_col_widths(self, ws, col_widths : dict, summary=False):
        '''iterates over {'A':30, 'B':40, 'C':35...} dict to resize worksheets' column widths. Summary ws wider columns with summary=True'''
        factor = 1.3 if summary else 1.05
        for col_letter in col_widths:
            adjusted_width = ((col_widths[col_letter] + 2) * factor)
            ws.column_dimensions[col_letter].width = adjusted_width
    
    def fill_format_summary(self):
        '''Forms a summary sheet report unpacks self.summary_table_obj to dynamic height table,
        change insert point of table with:
        REPORT_START_ROW, REPORT_START_COL'''
        self.s_ws = self.wb[SUMMARY_SHEET_NAME]
        self.col_widths = {}   
        self.row_cursor = REPORT_START_ROW
        self.__add_summary_headers()
        self.__color_table_headers()
        # Add data for each currency:
        for currency, date_objs in self.summary_table_obj.items():
            self.__apply_horizontal_line(self.row_cursor)
            self.s_ws.cell(self.row_cursor, REPORT_START_COL).value = currency
            self.s_ws.cell(self.row_cursor, REPORT_START_COL).font = BOLD_STYLE
            # Writing data to rest of columns:
            for date, date_orders in date_objs.items():
                self.s_ws.cell(self.row_cursor, REPORT_START_COL + 1).value = date
                self.__fill_format_date_data(date_orders)
                self.row_cursor += 1
        # Adjust column widths
        self._adjust_col_widths(self.s_ws, self.col_widths, summary=True)

    def __add_summary_headers(self):
        '''writes fixed headers in summary sheet'''
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 3).value = 'Daily Breakdown'
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 3).font = BOLD_STYLE
        self.row_cursor += 1
        for idx, header in enumerate(SUMMARY_HEADERS):
            self.s_ws.cell(self.row_cursor, REPORT_START_COL + idx).value = header
            self.__update_col_widths(REPORT_START_COL + idx, header, zero_indexed=False)
            self.s_ws.cell(self.row_cursor, REPORT_START_COL + idx).font = BOLD_STYLE
        self.row_cursor += 1

    def __color_table_headers(self):
        '''Colors range defined in generator in summary sheet'''
        for row, col in self.__header_cells_generator():
            self.s_ws.cell(row, col).fill = BACKGROUND_COLOR_STYLE
    
    @staticmethod
    def __header_cells_generator():
        '''generator for daily breakdown headers coloring, yields row, col'''
        for row in [REPORT_START_ROW, REPORT_START_ROW + 1]:
            for col in range(REPORT_START_COL, len(SUMMARY_HEADERS) + REPORT_START_COL):
                yield row, col

    def __apply_horizontal_line(self, row):
        '''adds horizonal line through A:H (c=1 case) in summary sheet at argument row top'''
        for col in range(REPORT_START_COL, len(SUMMARY_HEADERS) + REPORT_START_COL):
            self.s_ws.cell(row, col).border = openpyxl.styles.Border(top=THIN_BORDER)
    
    def __fill_format_date_data(self, date_orders : list):
        '''calculates neccessary data and fills, formats data in summary sheet in single row'''
        # Data does not update column widths, only headers. If data formats, scope were to change, function shall be updated 
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 2).value = self._get_segment_total(date_orders)
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 2).font = BOLD_STYLE
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 2).number_format = '#,##0.00'
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 3).value = len(date_orders)
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 3).font = BOLD_STYLE
        # Separating regions, filling data:
        VAT_orders = list(filter(lambda order: (order['item-tax'] > 0), date_orders))
        NON_VAT_orders = list(filter(lambda order: (order['item-tax'] <= 0), date_orders))
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 4).value = self._get_segment_total(VAT_orders)
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 4).number_format = '#,##0.00'
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 5).value = len(VAT_orders)
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 6).value = self._get_segment_total(NON_VAT_orders)
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 6).number_format = '#,##0.00'
        self.s_ws.cell(self.row_cursor, REPORT_START_COL + 7).value = len(NON_VAT_orders)

    @staticmethod
    def _get_segment_total(orders : list) -> float:
        '''Returns a sum of all orders' item-price + item-tax in list of order dicts'''
        total = 0
        for order in orders:
            total += order['item-price'] + order['item-tax']
        return round(total, 2)


    def export(self, wb_name : str):
        '''Creates workbook, and exports class objects: segments_orders_obj and summary_table_obj to
        segment worksheets and creates report summary sheet, saves new workbook'''
        self.wb = openpyxl.Workbook()
        ws = self.wb.active
        ws.title = SUMMARY_SHEET_NAME
        for segment, segment_orders in self.segments_orders_obj.items():
            self._data_to_sheet(segment, segment_orders)
        self.fill_format_summary()
        self.wb.save(wb_name)
        self.wb.close()


if __name__ == "__main__":
    pass