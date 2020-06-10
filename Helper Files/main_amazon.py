from amzn_parser_utils import get_output_dir, get_datetime_obj, alert_vba_date_count
from parse_orders import ParseOrders
from orders_db import OrdersDB
from datetime import datetime
import logging
import sys
import csv
import os

# DELETE LATER
from amzn_parser_utils import orders_column_to_file

# GLOBAL VARIABLES
TESTING = True
TEST_TODAY_DATE = '2020-04-17'
EXPECTED_SYS_ARGS = 2
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
VBA_KEYERROR_ALERT = 'ERROR_IN_SOURCE_HEADERS'
VBA_OK = 'EXPORTED_SUCCESSFULLY'
TEST_AMZN_EXPORT_TXT = r'/home/devyo/Coding/Git/Amazon Accounting Report/Amazon exports/run1.txt'

# Logging config:
log_path = os.path.join(get_output_dir(client_file=False), 'amazon_accounting.log')
logging.basicConfig(handlers=[logging.FileHandler(log_path, 'a', 'utf-8')], level=logging.DEBUG)


def get_cleaned_orders(source_file:str) -> list:
    '''returns cleaned orders (as cleaned in clean_orders func) from source_file arg path'''
    raw_orders = get_raw_orders(source_file)
    cleaned_orders = clean_orders(raw_orders)
    return cleaned_orders

def get_raw_orders(source_file:str) -> list:
    '''returns raw orders as list of dicts for each order in txt source_file'''
    with open(source_file, 'r', encoding='utf-8') as f:
        source_contents = csv.DictReader(f, delimiter='\t')
        raw_orders = [{header : value for header, value in row.items()} for row in source_contents]
    return raw_orders

def clean_orders(orders:list) -> list:
    '''replaces plus sign in phone numbers with 00'''
    for order in orders:
        try:
            order['buyer-phone-number'] = order['buyer-phone-number'].replace('+', '00')
        except KeyError as e:
            logging.warning(f'New header in source file. Alert VBA on new header. Error: {e}')
            print(VBA_KEYERROR_ALERT)
    return orders

def parse_export_orders(testing:bool, parse_orders:list, loaded_txt:str):
    '''interacts with classes (ParseOrders, OrdersDB) to filter new orders, export desired files and push new orders to db'''
    db_client = OrdersDB(parse_orders, loaded_txt)
    new_orders = db_client.get_new_orders_only()
    logging.info(f'After checking with database, further processing: {len(new_orders)} new orders')
    ParseOrders(new_orders, db_client).export_orders(testing)

def remove_todays_orders(orders:list) -> list :
    '''returns a list of orders dicts, whose purchase date up to, but not including today's date (deletes todays orders), alerts VBA'''
    try:
        today_date = datetime.fromisoformat(TEST_TODAY_DATE) if TESTING else datetime.today().date().isoformat()
        orders_until_today = list(filter(lambda order: get_datetime_obj(order['payments-date']) < today_date, orders))
        not_processing_count = len(orders) - len(orders_until_today)
        alert_vba_date_count(today_date, not_processing_count)
        logging.info(f'Orders passed today date filtering: {len(orders_until_today)}/{len(orders)}')
        return orders_until_today
    except Exception as e:
        logging.critical(f'Unknown error: {e} while filtering out todays orders. Date used: {today_date}; is valid datetime obj: {isinstance(today_date, datetime)}')
        print(VBA_ERROR_ALERT)
        sys.exit()

def parse_args():
    '''accepts txt_path from as command line argument'''
    if len(sys.argv) == EXPECTED_SYS_ARGS:
        txt_path = sys.argv[1]
        logging.info(f'Accepted sys args on launch: txt_path: {txt_path}')
        return txt_path
    else:
        print(VBA_ERROR_ALERT)
        logging.critical(f'Error parsing arguments on script initialization in cmd. Arguments provided: {len(sys.argv)} Expected: {EXPECTED_SYS_ARGS}')
        sys.exit()


def main(testing, amazon_export_txt_path):
    '''Main function executing parsing of provided txt file and outputing csv, xlsx files'''
    logging.info(f'\n NEW RUN STARTING: {datetime.today().strftime("%Y.%m.%d %H:%M")}')    
    if not testing:
        txt_path = parse_args()
    else:
        print('RUNNING IN TESTING MODE')
        txt_path = amazon_export_txt_path
    if os.path.exists(txt_path):
        logging.info('file exists, continuing to processing...')
        print(f'File {os.path.basename(txt_path)} exists')
        cleaned_source_orders = get_cleaned_orders(txt_path)
        orders_for_parsing = remove_todays_orders(cleaned_source_orders)
        parse_export_orders(testing, orders_for_parsing, txt_path)
        print(VBA_OK)
    else:
        logging.critical(f'Provided file {txt_path} does not exist.')
        print(VBA_ERROR_ALERT)
        sys.exit()
    logging.info(f'\nRUN ENDED: {datetime.today().strftime("%Y.%m.%d %H:%M")}\n')



if __name__ == "__main__":
    main(testing=TESTING, amazon_export_txt_path=TEST_AMZN_EXPORT_TXT)