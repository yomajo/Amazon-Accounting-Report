import logging
import os
from datetime import datetime
from collections import defaultdict
from accounting_utils import get_output_dir, get_EU_countries_from_txt, get_order_tax
from orders_report import AmazonEUOrdersReport, AmazonCOMOrdersReport
from constants import VBA_ERROR_ALERT, VBA_KEYERROR_ALERT, VBA_NO_NEW_JOB


# GLOBAL VARIABLES
EU_COUNTRIES_TXT = 'EU Countries.txt'


class ParseOrders():
    '''Input: orders as list of dicts, parses orders, groups, forms output object;
    passes to OrdersReport class which creates report in xlsx format.
    Interacts with database client instance; main method:
    
    export_orders(testing=False) : groups orders by EU/ non-EU orders, with nesting based on currency.    
    when testing flag = True, export is suspended, but orders passed to class are still added to database
    
    Args:
    - orders : list - list of order dictionaries
    - sales_channel : str - 'AmazonEU'/'AmazonCOM'/'Amazon Warehouse' to differenciate different report
    - db_client:object - db client to iteract with during program runtime'''
    
    def __init__(self, all_orders: list, db_client: object, sales_channel: str, proxy_keys: dict):
        self.all_orders = all_orders
        self.db_client = db_client
        self.sales_channel = sales_channel
        self.proxy_keys = proxy_keys
        self.eu_orders = []
        self.non_eu_orders = []
    
    def _prepare_filepaths(self):
        '''creates cls variables of files abs paths to be created one dir above this script dir'''
        output_dir = get_output_dir()
        date_stamp = datetime.today().strftime("%Y.%m.%d %H.%M")
        self.report_path = os.path.join(output_dir, f'{self.sales_channel} Report {date_stamp}.xlsx')
    
    def split_orders_by_region(self):
        '''Sorts all orders into eu/non_eu regions based ship country and sales channel'''
        self.eu_countries = self._get_EU_countries_list_from_file()
        for order in self.all_orders:
            try:
                if order[self.proxy_keys['ship-country']] in self.eu_countries:
                    # Add EU orders with tax = 0 to non-vat (non-eu)
                    if self.sales_channel == 'AmazonEU' and get_order_tax(order, self.proxy_keys) == 0:
                        self.non_eu_orders.append(order)
                    else:
                        self.eu_orders.append(order)
                else:
                    self.non_eu_orders.append(order)
            except KeyError:
                logging.exception(f'Could not find item-tax in (using proxy keys) order keys. Order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_KEYERROR_ALERT)
                exit()
            except ValueError:
                logging.exception(f'Could not return float value for item-tax (using proxy keys) in order: {order}\nClosing connection to database, alerting VBA, exiting...')
                self.db_client.close_connection()
                print(VBA_ERROR_ALERT)
                exit()
    
    def _get_EU_countries_list_from_file(self):
        '''returns list of EU member countries from TXT file'''
        current_dir = get_output_dir(client_file=False)
        txt_abspath = os.path.join(current_dir, EU_COUNTRIES_TXT)
        logging.debug(f'Trying to access EU countries txt file: {txt_abspath}')
        return get_EU_countries_from_txt(txt_abspath)

    def exit_no_new_orders(self):
        '''terminate if all lists after sorting are empty'''
        if not self.eu_orders and not self.non_eu_orders:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            exit()

    def prepare_export_obj(self) -> dict:
        '''Constructs a dict data object for OrdersReport class. Output format:
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
                                }'''
        eu_currency_grouped = self.get_region_currency_based_dict(self.eu_orders)
        non_eu_currency_grouped = self.get_region_currency_based_dict(self.non_eu_orders)
        self.export_obj = {'EU' : eu_currency_grouped, 'NON-EU' : non_eu_currency_grouped}
        logging.debug(f'Returning export object with keys: {self.export_obj.keys()}')
        return self.export_obj

    def get_region_currency_based_dict(self, region_orders: list) -> dict:
        '''returns currency grouped dict.
        Example: {'EUR': [order1, order2...], 'USD':[order1, order2...], ...}'''
        currency_based_dict = defaultdict(list)
        for order in region_orders:
            order_currency = order[self.proxy_keys['currency']].upper()
            currency_based_dict[order_currency].append(order)
        return currency_based_dict

    def export_report(self):
        '''creates AmazonEUOrdersReport or AmazonCOMOrdersReport instance, and exports report in xlsx format'''
        try:
            if self.sales_channel in ['AmazonEU', 'Amazon Warehouse']:
                logging.info(f'Passing orders to create report with {AmazonEUOrdersReport.__name__} class')
                AmazonEUOrdersReport(self.export_obj, self.eu_countries, self.sales_channel, self.proxy_keys).export(self.report_path)
            elif self.sales_channel == 'AmazonCOM':
                logging.info(f'Passing orders to create report with {AmazonCOMOrdersReport.__name__} class')
                AmazonCOMOrdersReport(self.export_obj, self.eu_countries, self.sales_channel, self.proxy_keys).export(self.report_path)
            logging.info(f'XLSX report {os.path.basename(self.report_path)} successfully created.')
        except:
            logging.exception(f'Unexpected error creating report. Closing database connection, alerting VBA, exiting ParseOrders...')
            self.db_client.close_connection()
            print(VBA_ERROR_ALERT)
            exit()
    
    def push_orders_to_db(self):
        '''adds all orders in this class to orders table in db'''
        count_added_to_db = self.db_client.add_orders_to_db()
        logging.info(f'Total of {count_added_to_db} new orders have been added to database, after exports were completed')

    def export_orders(self, testing=False):
        '''Summing up tasks inside ParseOrders class'''
        self._prepare_filepaths()
        self.split_orders_by_region()
        self.exit_no_new_orders()
        self.prepare_export_obj()
        if testing:
            logging.info(f'Running in testing {testing} environment. Change behaviour in export_orders method in ParseOrders class')
            print(f'Running in testing {testing} environment. Change behaviour in export_orders method in ParseOrders class')
            print('ENABLED REPORT EXPORT WHILE TESTING')            
            self.export_report()
            self.push_orders_to_db()
            return
        self.export_report()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass