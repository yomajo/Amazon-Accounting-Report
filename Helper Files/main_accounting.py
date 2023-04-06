import logging
import sys
import csv
import os
from datetime import datetime
import sqlalchemy.sql.default_comparator    #neccessary for executable packing
from accounting_utils import get_output_dir, get_datetime_obj, alert_vba_date_count
from accounting_utils import get_file_encoding_delimiter, delete_file, dump_to_json, orders_column_to_file
from parse_orders import ParseOrders
from orders_db import SQLAlchemyOrdersDB
from constants import SALES_CHANNEL_PROXY_KEYS, VBA_ERROR_ALERT, VBA_KEYERROR_ALERT, VBA_OK, VBA_COUNTRYLESS_ALERT


TEST_CASES = [
    {'channel': 'AmazonEU', 'file': r'C:\Coding\Ebay\Working\Backups\Amazon exports\AmazonEU 2023.04.06.txt'},
    {'channel': 'AmazonCOM', 'file': r'C:\Coding\Ebay\Working\Backups\Amazon exports\COM 2022.10.04.txt'},
    {'channel': 'Amazon Warehouse', 'file': r'C:\Coding\Ebay\Working\Backups\Amazon warehouse csv\warehouse2.csv'},
]

# GLOBAL VARIABLES
TESTING = False
TEST_CASE = TEST_CASES[0]
# TEST_TODAY_DATE = '2022-06-27'
TEST_TODAY_DATE = datetime.now().strftime('%Y-%m-%d')   # Hardcode in format: '2021-08-19' if needed when testing

SALES_CHANNEL = TEST_CASE['channel']
ORDERS_SOURCE_FILE = TEST_CASE['file']
EXPECTED_SYS_ARGS = 3

# Logging config:
log_path = os.path.join(get_output_dir(client_file=False), 'report.log')
logging.basicConfig(handlers=[logging.FileHandler(log_path, 'a', 'utf-8')], level=logging.INFO)


def get_cleaned_orders(source_file:str, sales_channel:str, proxy_keys:dict) -> list:
    '''returns cleaned orders (as cleaned in clean_orders func) from source_file arg path'''
    encoding, delimiter = get_file_encoding_delimiter(source_file)
    logging.info(f'{os.path.basename(source_file)} detected encoding: {encoding}, delimiter <{delimiter}>')
    raw_orders = get_raw_orders(source_file, encoding, delimiter)
    logging.info(f'Loaded {os.path.basename(source_file)} has {len(raw_orders)} raw orders. Filtering out todays orders...')
    cleaned_orders = remove_todays_orders(raw_orders, sales_channel, proxy_keys)
    if TESTING:
        replace_old_testing_json(raw_orders, 'DEBUG_raw_all.json')
        replace_old_testing_json(cleaned_orders, 'DEBUG_filtred_todays.json')
    return cleaned_orders

def get_raw_orders(source_file:str, encoding:str, delimiter:str) -> list:
    '''returns raw orders as list of dicts for each order in txt source_file'''
    with open(source_file, 'r', encoding=encoding) as f:
        source_contents = csv.DictReader(f, delimiter=delimiter)
        raw_orders = [{header : value for header, value in row.items()} for row in source_contents]
    return raw_orders

def replace_old_testing_json(raw_orders, json_fname:str):
    '''deletes old json, exports raw orders to json file'''
    output_dir = get_output_dir(client_file=False)
    json_path = os.path.join(output_dir, json_fname)
    delete_file(json_path)
    dump_to_json(raw_orders, json_fname)

def remove_todays_orders(orders: list, sales_channel: str, proxy_keys: dict) -> list:
    '''returns a list of orders dicts, whose purchase date up to, but not including today's date (deletes todays orders), alerts VBA'''
    try:
        today_date = get_today_obj()
        # for VBA, logging, constructing str representation
        today_str = today_date.strftime('%Y-%m-%d')
        
        logging.info(f'Filter date used in program: {today_date}. Passing to vba and logging strftime format: {today_str}')
        orders_until_today = list(filter(lambda order: get_datetime_obj(order[proxy_keys['payments-date']], sales_channel) < today_date, orders))
        not_processing_count = len(orders) - len(orders_until_today)

        alert_vba_date_count(today_str, not_processing_count)
        logging.info(f'Orders passed today date filtering: {len(orders_until_today)}/{len(orders)}')
        return orders_until_today
    except KeyError as e:
        logging.critical(f'Err: {e} in remove_todays_orders method. Probable key not found: payments-date, (sales channel: {sales_channel})')
        print(VBA_KEYERROR_ALERT)
        exit()
    except Exception as e:
        logging.critical(f'Unknown error: {e} while filtering out todays orders. Date used: {today_date}; sales_channel: {sales_channel}')
        print(VBA_ERROR_ALERT)
        exit()

def get_today_obj(): 
    '''returns instance of datetime library corresponding to date (no time) for today used in rest of program'''
    if TESTING:
        return datetime.strptime(TEST_TODAY_DATE, '%Y-%m-%d')
    else:
        dt_today_date_only = datetime.today().strftime('%Y-%m-%d')
        return datetime.strptime(dt_today_date_only, '%Y-%m-%d')

def remove_countryless(orders: list, proxy_keys: dict) -> list:
    '''removes orders w/o defined country, alerts VBA, exports IDs to txt file if present'''
    logging.debug(f'Before countryless filter: {len(orders)} orders')
    countryless = list(filter(lambda x: x[proxy_keys['ship-country']] == '', orders))
    if countryless:
        logging.info(f'Removed {len(countryless)} country-less orders')
        fpath = orders_column_to_file(countryless, proxy_keys['secondary-order-id'])
        print(VBA_COUNTRYLESS_ALERT)
        logging.info(f'Country-less orders have been exported to txt {fpath} file, VBA alerted. Proceeding...')
    return list(filter(lambda x: x[proxy_keys['ship-country']] != '', orders))

def parse_args():
    '''returns source_fpath, sales_channel from cli args or hardcoded testing variables'''    
    if TESTING:
        print(f'--- RUNNING IN TESTING MODE. Using hardcoded args ch: {SALES_CHANNEL}, f: {os.path.basename(ORDERS_SOURCE_FILE)}---')
        logging.warning('--- RUNNING IN TESTING MODE. Using hardcoded args---')
        assert SALES_CHANNEL in SALES_CHANNEL_PROXY_KEYS.keys(), f'Unexpected sales_channel value passed from VBA side: {SALES_CHANNEL}'
        return ORDERS_SOURCE_FILE, SALES_CHANNEL
    try:
        assert len(sys.argv) == EXPECTED_SYS_ARGS, 'Unexpected number of sys.args passed. Check TESTING mode'
        source_fpath = sys.argv[1]
        sales_channel = sys.argv[2]
        logging.info(f'Accepted sys args on launch: source_fpath: {source_fpath}; sales_channel: {sales_channel}. Whole sys.argv: {list(sys.argv)}')
        assert sales_channel in SALES_CHANNEL_PROXY_KEYS.keys(), f'Unexpected sales_channel value passed from VBA side: {sales_channel}'
        return source_fpath, sales_channel
    except Exception as e:
        print(VBA_ERROR_ALERT)
        logging.critical(f'Error parsing arguments on script initialization in cmd. Arguments provided: {list(sys.argv)} Number Expected: {EXPECTED_SYS_ARGS}. Err: {e}')
        exit()

def main():
    '''Main function executing parsing of provided txt file and outputing csv, xlsx files'''    
    logging.info(f'\n NEW RUN STARTING: {datetime.today().strftime("%Y.%m.%d %H:%M")}')    
    source_fpath, sales_channel = parse_args()
    proxy_keys = SALES_CHANNEL_PROXY_KEYS[sales_channel]
    logging.debug(f'Loading file: {os.path.basename(source_fpath)}. Using proxy keys matching key: {sales_channel} in SALES_CHANNEL_PROXY_KEYS')

    # Get cleaned (filter out today's orders) source orders
    cleaned_source_orders = get_cleaned_orders(source_fpath, sales_channel, proxy_keys)

    # dont store / evaluate country-less orders
    valid_orders = remove_countryless(cleaned_source_orders, proxy_keys)

    db_client = SQLAlchemyOrdersDB(valid_orders, source_fpath, sales_channel, proxy_keys, testing=TESTING)
    new_orders = db_client.get_new_orders_only()
    logging.info(f'Loaded file contains: {len(cleaned_source_orders)} (b4 {TEST_TODAY_DATE} date and countryless filters. Further processing: {len(new_orders)} orders')

    # Parse orders, export target files
    ParseOrders(new_orders, db_client, sales_channel, proxy_keys).export_orders(TESTING)
    print(VBA_OK)
    logging.info(f'\nRUN ENDED: {datetime.today().strftime("%Y.%m.%d %H:%M")}\n')


if __name__ == "__main__":
    main()