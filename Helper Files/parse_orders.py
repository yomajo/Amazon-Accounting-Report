from amzn_parser_utils import get_output_dir
from orders_report import OrdersReport
from collections import defaultdict
from datetime import datetime
import logging
import sys
import csv
import os


# GLOBAL VARIABLES
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'
VBA_NO_NEW_JOB = 'NO NEW JOB'
VBA_KEYERROR_ALERT = 'ERROR_IN_SOURCE_HEADERS'
DPOST_REF_CHARLIMIT_PER_CELL = 28


class ParseOrders():
    '''Input: orders as list of dicts, parses orders, groups, forms output object;
    passes to OrdersReport class which creates report in xlsx format.
    Interacts with database client instance; main method:
    
    export_orders(testing=False) : groups orders by EU/ non-EU orders, with nesting based on currency.    
    when testing flag = True, export is suspended, but orders passed to class are still added to database'''
    
    def __init__(self, all_orders : list, db_client : object):
        self.all_orders = all_orders
        self.db_client = db_client
        self.eu_orders = []
        self.non_eu_orders = []
    
    def _prepare_filepaths(self):
        '''creates cls variables of files abs paths to be created one dir above this script dir'''
        output_dir = get_output_dir()
        date_stamp = datetime.today().strftime("%Y.%m.%d %H.%M")
        self.report_path = os.path.join(output_dir, f'Amazon Orders Report {date_stamp}.xlsx')
    
    def split_orders_by_tax_region(self):
        '''splits all_orders into two lists: EU (VAT (item-tax) > 0) and NON-EU (VAT = 0); Exits if resulting lists are empty'''    
        for order in self.all_orders:
            try:
                if float(order['item-tax']) > 0:
                    self.eu_orders.append(order)
                else:
                    self.non_eu_orders.append(order)
            except KeyError:
                logging.exception(f'Could not find item-tax in order keys. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_KEYERROR_ALERT)
                sys.exit()                
            except ValueError:
                logging.exception(f'Could not return float value for item-tax in order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_ERROR_ALERT)
                sys.exit()
        self.exit_no_new_orders()
    
    def exit_no_new_orders(self):
        if not self.eu_orders and not self.non_eu_orders:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            sys.exit()

    def prepare_export_obj(self):
        '''constructs a dict data object for OrdersReport class. Output format:
        export_data = {
                eu_orders: {
                            currency1: [order1, order2, order...],
                            currency2: [order1, order2, order...],
                            currency_n : [order1, order2, order...]
                            },
                non_eu_orders: {
                            currency1: [order1, order2, order...],
                            currency2: [order1, order2, order...],
                            currency_n : [order1, order2, order...]
                                }
                    }'''
        eu_currency_grouped = self.get_region_currency_based_dict(self.eu_orders)
        non_eu_currency_grouped = self.get_region_currency_based_dict(self.non_eu_orders)
        self.export_obj = {'EU' : eu_currency_grouped, 'NON-EU' : non_eu_currency_grouped}
        return self.export_obj

    @staticmethod
    def get_region_currency_based_dict(region_orders : list) -> dict:
        currency_based_dict = defaultdict(list)
        for order in region_orders:
            currency_based_dict[order['currency']].append(order)
        return currency_based_dict

    def export_report(self):
        '''creates OrdersReport instance, and exports report in xlsx format'''
        try:
            OrdersReport(self.export_obj).export(self.report_path)
            logging.info(f'XLSX report {os.path.basename(self.report_path)} successfully created.')
        except:
            logging.exception(f'Unexpected error creating report. Closing database connection, alerting VBA, exiting...')
            self.db_client.close_connection()
            print(VBA_ERROR_ALERT)
            sys.exit()
    
    def push_orders_to_db(self):
        '''adds all orders in this class to orders table in db'''
        count_added_to_db = self.db_client.add_orders_to_db()
        logging.info(f'Total of {count_added_to_db} new orders have been added to database, after exports were completed')

    def export_orders(self, testing=False):
        '''Summing up tasks inside ParseOrders class'''
        self._prepare_filepaths()
        self.split_orders_by_tax_region()
        self.prepare_export_obj()        
        if testing:
            logging.info(f'Due to flag testing value: {testing}. Order export and adding to database suspended. Change behaviour in export_orders method in ParseOrders class')
            print(f'Due to flag testing value: {testing}. Order export and adding to database suspended. Change behaviour in export_orders method in ParseOrders class')
            self.db_client.close_connection()
            print('ENABLED REPORT EXPORT WHILE TESTING')
            self.export_report()
            # self.push_orders_to_db()
            return
        self.export_report()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass