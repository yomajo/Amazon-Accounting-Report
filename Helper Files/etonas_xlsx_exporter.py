from amzn_parser_constants import ETONAS_HEADERS, ETONAS_HEADERS_MAPPING
from amzn_parser_utils import get_product_category
import logging
import openpyxl
import sys

# GLOBAL VARIABLES
ETONAS_CHARLIMIT_PER_CELL = 32
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
VBA_ETONAS_CHARTLIMIT_ALERT = 'ETONAS_CHARLIMIT_WARNING'


class EtonasExporter():
    '''accepts etonas orders data as list of dicts, creates formatted xlsx output'''
    def __init__(self, etonas_orders : list, etonas_path : str):
        self.etonas_orders = etonas_orders
        self.etonas_path = etonas_path
    
    def reformat_data_to_etonas_output(self, orders_data : list):
        '''reduces input data to that needed in output csv'''
        try:
            etonas_ready_data = []
            for order_dict in self.etonas_orders:
                reduced_order_dict = self.prepare_etonas_order_contents(order_dict)
                etonas_ready_data.append(reduced_order_dict)
            return etonas_ready_data
        except Exception as e:
            print(VBA_ERROR_ALERT)
            logging.warning(f'Error while iterating collected row dicts and trying to reduce in ETONAS. Error: {e}')

    def prepare_etonas_order_contents(self, order_dict : dict):
        d_with_output_keys = {}
        first_name, last_name = self.get_fname_lname(order_dict)
        for header in ETONAS_HEADERS:
            if header in ETONAS_HEADERS_MAPPING.keys():
                d_with_output_keys[header] = order_dict[ETONAS_HEADERS_MAPPING[header]]
                # warn in VBA if char limit per cell is exceeded in Etonas address lines 1/2/3/4
                if 'address' in header.lower() and len(d_with_output_keys[header]) > ETONAS_CHARLIMIT_PER_CELL:
                    logging.info(f'Order with key {header} and value {d_with_output_keys[header]} triggered VBA warning for charlimit set by Etonas')
                    print(VBA_ETONAS_CHARTLIMIT_ALERT)
            elif header == 'First_name':
                d_with_output_keys['First_name'] = first_name
            elif header == 'Last_name':
                d_with_output_keys['Last_name'] = last_name
            elif header == 'Contents':
                d_with_output_keys['Contents'] = get_product_category(order_dict['product-name'])
            else:
                d_with_output_keys[header] = ''
        return d_with_output_keys
    
    @staticmethod
    def get_fname_lname(order : dict):
        '''returns first and last name split at space in order dict key ['recipient-name']'''
        try:
            f_name, l_name = order['recipient-name'].split(' ', 1)
            return f_name, l_name
        except KeyError as e:
            logging.critical(f'No recipient-name key. Error: {e} Order: {order}')
            print(VBA_ERROR_ALERT)
            sys.exit()
        except ValueError as e:
            logging.info(f'recipient-name seems to be one word, without a need to unpack. Error: {e}. recipient-name: {order["recipient-name"]}. Returning empty last_name')
            return order['recipient-name'], ''

    @staticmethod
    def _write_headers(worksheet, headers):
        for col, header in enumerate(headers, 1):
            worksheet.cell(1, col).value = header    

    @staticmethod
    def range_generator(orders_data, headers):
        for row, _ in enumerate(orders_data):
            for col, _ in enumerate(headers):
                yield row, col
    
    def _write_etonas_orders(self, worksheet, headers, orders_data):
        for row, col in self.range_generator(orders_data, headers):
            working_dict = orders_data[row]
            key_pointer = headers[col]
            # offsets due to excel vs python numbering  + headers in row 1
            worksheet.cell(row + 2, col + 1).value = working_dict[key_pointer]

    def adjust_col_widths(self, worksheet):
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.1
            worksheet.column_dimensions[col_letter].width = adjusted_width

    def export(self):
        reheaded_etonas_orders = self.reformat_data_to_etonas_output(self.etonas_orders)
        wb = openpyxl.Workbook()
        ws = wb.active
        self._write_headers(ws, ETONAS_HEADERS)
        self._write_etonas_orders(ws, ETONAS_HEADERS, reheaded_etonas_orders)        
        self.adjust_col_widths(ws)
        wb.save(self.etonas_path)


if __name__ == "__main__":
    pass