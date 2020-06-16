from amzn_parser_utils import simplify_date, col_to_letter
from amzn_parser_constants import TEMPLATE_SHEET_MAPPING
from collections import defaultdict
import itertools
import openpyxl
import os

# GLOBAL VARIABLES
SUMMARY_SHEET_NAME = 'Summary'
D_BREAKDOWN_HEADERS = ['Segment', 'Date', 'Total', 'Orders', 'Average']
BOLD_STYLE = openpyxl.styles.Font(bold=True, name='Calibri')
BACKGROUND_COLOR_STYLE = openpyxl.styles.PatternFill(fgColor='D6DEFF',fill_type='solid')
THIN_BORDER = openpyxl.styles.Side(border_style='thin')
SHEET_HEADERS = list(TEMPLATE_SHEET_MAPPING.keys())


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
        self.export_obj = export_obj
        self.__get_report_objs()
    
    def __get_report_objs(self):
        '''prepares cls variables from passed raw data - export_obj:
         1. self.segments_orders_obj, 2. self.daily_breakdown_obj, 3. self.totals'''
        self.segments_orders_obj = self._get_segments_orders_obj(self.export_obj)
        self.daily_breakdown_obj = self._get_segments_daily_breakdown_obj(self.segments_orders_obj)
        self.totals = self._get_totals_counts_avgs(self.segments_orders_obj)
    
    def _get_segments_orders_obj(self, export_obj : dict) -> dict:
        '''returns dict of dicts for each segment (sheets) and corresponding list of orders
        returned object: {'EU EUR' : [order1, order2, ...], 'NON-EU GBP' : [order11, order 22...], ...}'''
        segments_orders = {}
        for region, currency in self._unpack_export_obj(export_obj):
            segments_orders[f'{region} {currency}'] = self._correct_orders_numbers_dates(export_obj[region][currency])
        return segments_orders

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

    def _get_totals_counts_avgs(self, segments_obj : dict) -> dict:
        '''Calulates and returns totals obj for each segment. Return format:
        {'EU EUR': {'Total':2066, 'Count':69, 'Average':29.94}, 'NON-EU GBP': {'Total':1534.72, 'Count':122, 'Average':12.58}, ...}'''
        segments_totals = {}
        for segment, segment_orders in segments_obj.items():
            s_total = self._get_segment_total(segment_orders)
            s_count = len(segment_orders)
            s_avg = round(s_total / s_count, 2)
            segments_totals[segment] = {'Total' : s_total, 'Orders' : s_count, 'Average' : s_avg}
        return segments_totals

    @staticmethod
    def _get_segment_total(orders : list) -> float:
        '''Returns a sum of all orders' item-price + item-tax in list of order dicts'''
        total = 0
        for order in orders:
            total += order['item-price'] + order['item-tax']
        return round(total, 2)

    def _get_segments_daily_breakdown_obj(self, segments_obj : dict) -> dict:
        '''returns object dissected by payment date stats for each segment. Example:
        {'EU EUR': {'date1': {'Total': X, 'Orders': Y, 'Average': Z}, 'date2': {'Total': X, 'Orders': Y, 'Average': Z}, ...},
        'NON-EU GBP': {'date1': {'Total': X, 'Orders': Y, 'Average': Z}, 'date2': {'Total': X, 'Orders': Y, 'Average': Z}, ...}, ...}'''
        daily_breakdown_obj = {}
        for segment, segment_orders in segments_obj.items():
            segment_dates_obj = self._get_payment_date_obj(segment_orders)
            # Note: stats by date keys, contrary to docstring, here returns obj, whose keys are payments dates
            stats_by_date = self._get_totals_counts_avgs(segment_dates_obj)
            daily_breakdown_obj[segment] = stats_by_date        
        return daily_breakdown_obj

    @staticmethod
    def _get_payment_date_obj(segment_orders : list) -> dict:
        '''forms a dictionary based on segment's payment dates as keys and orders list as values.
        Output format: {
            {'YYYY-MM-D1':[order1, order2, ...]},
            {'YYYY-MM-D2':[order1, order2, ...]},
            ...}'''
        payment_date_orders = defaultdict(list)
        for order in segment_orders:
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
            self._update_col_widths(col, header)
            ws.cell(1, col + 1).value = header

    @staticmethod
    def range_generator(orders_data, headers):
        for row, _ in enumerate(orders_data):
            for col, _ in enumerate(headers):
                yield row, col
    
    def _write_sheet_orders(self, ws, orders_data : list):
        for row, col in self.range_generator(orders_data, SHEET_HEADERS):
            working_dict = orders_data[row]
            key_pointer = TEMPLATE_SHEET_MAPPING[SHEET_HEADERS[col]]
            self._update_col_widths(col, str(working_dict[key_pointer]))
            # offsets due to excel vs python numbering  + headers in row 1
            ws.cell(row + 2, col + 1).value = working_dict[key_pointer]

    def _update_col_widths(self, col : int, cell_value : str, zero_indexed=True):
        '''runs on each cell. Forms a dictionary {'A':30, 'B':15...} for max column widths in worksheet (width as length of max cell)'''
        col_letter = col_to_letter(col, zero_indexed=zero_indexed)
        if col_letter in self.col_widths:
            # check for length, update if current cell length exceeds current entry for column
            if len(cell_value) > self.col_widths[col_letter]:
                self.col_widths[col_letter] = len(cell_value)
        else:
            self.col_widths[col_letter] = len(cell_value)

    def _adjust_col_widths(self, ws, col_widths : dict, summary=False):
        '''iterates over {'A':30, 'B':40, 'C':35...} dict to resize worksheets' column widths. Summary ws wider columns with summary=True'''
        factor = 1.3 if summary else 1.05
        for col_letter in col_widths:
            adjusted_width = ((col_widths[col_letter] + 2) * factor)
            ws.column_dimensions[col_letter].width = adjusted_width
    
    def fill_summary(self):
        '''Forms a summary statistics sheet / report'''
        active_ws = self.wb[SUMMARY_SHEET_NAME]
        self.col_widths = {}
        self.write_totals_table(active_ws)
        self.write_daily_breakdown_table(active_ws)
        self._adjust_col_widths(active_ws, self.col_widths, summary=True)
        self.color_table_headers(active_ws)
        
    def write_totals_table(self, ws):
        '''fills out first table in main sheet of report.
        Table height is fixed, but width is flexible depending on data segments'''
        ws.cell(1, 3).value = 'Totals'
        ws.cell(1, 3).font = BOLD_STYLE
        # Write vertical headers (Total, Orders, Average in A3:A5)
        any_values_dict = list(self.totals.values())[0]
        for row, values_header in enumerate(any_values_dict.keys()):
            self._update_col_widths(1, values_header)
            ws.cell(row + 3, 1).value = values_header
        # Fill segment name and values (range B2:[n]5):
        for col, (segment, stats) in enumerate(self.totals.items()):
            self._update_col_widths(col + 2, segment, zero_indexed=False)
            ws.cell(2, col + 2).value = segment
            ws.cell(2, col + 2).font = BOLD_STYLE
            # Skipping update col_widths. Numbers unlikely to break out beyond header titles
            ws.cell(3, col + 2).value = stats['Total']
            ws.cell(3, col + 2).number_format = '#,##0.00'
            ws.cell(4, col + 2).value = stats['Orders']
            ws.cell(5, col + 2).value = stats['Average']
            ws.cell(5, col + 2).number_format = '#,##0.00'

    def write_daily_breakdown_table(self, ws):
        '''fills out second table, which breaks down daily statistics for each segment'''
        self.row_cursor = 8
        ws.cell(self.row_cursor - 1, 3).value = 'Daily Breakdown'
        ws.cell(self.row_cursor - 1, 3).font = BOLD_STYLE
        self._write_d_breakdown_headers(ws)
        self._write_d_breakdown_data(ws)

    def _write_d_breakdown_headers(self, ws):
        '''writes fixed headers for second table in summary sheet'''
        for col, d_breakdown_header in enumerate(D_BREAKDOWN_HEADERS):
            self._update_col_widths(col + 1, d_breakdown_header)
            ws.cell(self.row_cursor, col + 1).value = d_breakdown_header
            ws.cell(self.row_cursor, col + 1).font = BOLD_STYLE
        self.row_cursor += 1

    def _write_d_breakdown_data(self, ws):
        '''iterates self.daily_breakdown_obj and fills second table in summary sheet with data'''
        for region, daily_breakdown in self.daily_breakdown_obj.items():
            # Entering and writing segment
            self._update_col_widths(1, region, zero_indexed=False)
            ws.cell(self.row_cursor, 1).value = region
            self.__apply_horizontal_line(ws, self.row_cursor)            
            for date, stats in daily_breakdown.items():
                # Entering and writing daily breakdown stats
                self._update_col_widths(2, date, zero_indexed=False)
                ws.cell(self.row_cursor, 2).value = date
                self.__write_date_stats(ws, stats, self.row_cursor)
                self.row_cursor += 1
    
    @staticmethod
    def __write_date_stats(ws, stats : dict, row : int):
        '''writes Total, Orders and Average fields in Summary Daily Breakdown table. Takes args: ws (worksheet obj);
        stats in format: {'Total':X, 'Orders':Y, 'Average':Z}''' 
        ws.cell(row, 3).value = stats['Total']
        ws.cell(row, 3).number_format = '#,##0.00'
        ws.cell(row, 4).value = stats['Orders']
        ws.cell(row, 5).value = stats['Average']
        ws.cell(row, 5).number_format = '#,##0.00'

    def __apply_horizontal_line(self, ws, row):
        '''adds horizonal line through A:E at given row top'''
        for col in range(1, 6):
            ws.cell(row, col).border = openpyxl.styles.Border(top=THIN_BORDER)
    
    def color_table_headers(self, ws):
        '''Colors fixed ranges defined in generator {}'''
        for row, col in itertools.chain(self.color_headers_table1(), self.color_headers_table2()):
            ws.cell(row, col).fill = BACKGROUND_COLOR_STYLE
    
    def color_headers_table1(self):
        '''generator for Totals headers coloring (adjustable width dependent on segments count)'''
        color_cols =  len(list(self.totals.keys())) + 2
        for row in [1, 2]:
            for col in range(1, color_cols):
                yield row, col
    
    def color_headers_table2(self):
        '''generator for daily breakdown headers coloring'''
        for row in [7, 8]:
            for col in range(1, 6):
                yield row, col


    def export(self, wb_name : str):
        '''Creates workbook, and exports class objects: segments_orders_obj, daily_breakdown_obj, totals to
        segment worksheets and creates report summary sheet, saves new workbook'''
        self.wb = openpyxl.Workbook()
        ws = self.wb.active
        ws.title = SUMMARY_SHEET_NAME
        for segment, segment_orders in self.segments_orders_obj.items():
            self._data_to_sheet(segment, segment_orders)
        self.fill_summary()
        self.wb.save(wb_name)
        self.wb.close()


if __name__ == "__main__":
    pass