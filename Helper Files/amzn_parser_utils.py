from datetime import datetime
import platform
import logging
import sys
import os

import openpyxl

# GLOBAL VARIABLES
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'


def get_level_up_abspath(absdir_path):
    '''returns directory absolute path one level up from passed abs path'''
    return os.path.dirname(absdir_path)

def get_total_price(order_dict : dict):
    '''returns a sum of 'item-price' and 'shipping-price' for given order'''
    try:
        item_price = order_dict['item-price']
        shipping_price = order_dict['shipping-price']
        return str(float(item_price) + float(shipping_price))
    except KeyError as e:
        logging.critical(f'Could not find item-price or shipping-price keys in provided dict: {order_dict} Error: {e}')
        print(VBA_ERROR_ALERT)
        sys.exit()
    except ValueError as e:
        logging.critical(f"Could not convert item-price or shipping-price to float. Both values: {order_dict['item-price']}; {order_dict['shipping-price']} Error: {e}")
        print(VBA_ERROR_ALERT)
        sys.exit()

def get_output_dir(client_file=True):
    '''returns target dir for output files depending on execution type (.exe/.py) and file type (client/systemic)'''
    # pyinstaller sets 'frozen' attr to sys module when compiling
    if getattr(sys, 'frozen', False):
        curr_folder = os.path.dirname(sys.executable)
    else:
        curr_folder = os.path.dirname(os.path.abspath(__file__))
    return get_level_up_abspath(curr_folder) if client_file else curr_folder

def file_to_binary(abs_fpath:str):
    '''returns binary data for file'''
    try:
        with open(abs_fpath, 'rb') as f:
            bfile = f.read()
        return bfile
    except FileNotFoundError as e:
        print(f'file_to_binary func got arg: {abs_fpath}; resulting in error: {e}')
        return None

def recreate_txt_file(abs_fpath:str, binary_data):
    '''outputs a file from given binary data'''
    try:
        with open(abs_fpath, 'wb') as f:
            f.write(binary_data)
    except TypeError:
        print(f'Expected binary when writing contents to file {abs_fpath}')

def is_windows_machine() -> bool:
    '''returns True if machine executing the code is Windows based'''
    machine_os = platform.system()
    return True if machine_os == 'Windows' else False

def orders_column_to_file(orders:list, dict_key:str):
    '''exports a column values of each orders list item for passed dict_key'''
    try:
        export_data = [order[dict_key] for order in orders]
        with open(f'export {dict_key}.txt', 'w', encoding='utf-8', newline='\n') as f:
            f.writelines('\n'.join(export_data))
        print(f'Data exported to: {os.path.dirname(os.path.abspath(__file__))} folder')
    except KeyError:
        print(f'Provided {dict_key} does not exist in passed orders list of dicts')

def alert_vba_date_count(filter_date, orders_count):
    '''Passing two variables for VBA to display for user in message box'''
    print(f'FILTER_DATE_USED: {filter_date}')
    print(f'SKIPPING_ORDERS_COUNT: {orders_count}')

def dkey_to_float(order_dict : dict, key_title : str) -> float:
    '''returns float value of order_dict[key_title]'''
    return float(order_dict[key_title])

def get_datetime_obj(date_str):
    '''returns tz-naive datetime obj from date string. Designed to work with str format: 2020-04-16T10:07:16+00:00'''
    try:
        return datetime.fromisoformat(date_str).replace(tzinfo=None)
    except ValueError:
        # Attempt to handle wrong/new date format here
        logging.warning(f'Change in format detected! Previous format: 2020-04-16T10:07:16+00:00. Current: {date_str} Attempting to parse string...')
        try:
            date_str_split = date_str.split('T')[0]
            return datetime.fromisoformat(date_str_split)
        except ValueError:
            logging.critical(f'Unable to create datetime from date string: {date_str}. Terminating.')
            print(VBA_ERROR_ALERT)
            sys.exit()


if __name__ == "__main__":
    pass