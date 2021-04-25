from amzn_parser_utils import get_output_dir, get_EU_countries_from_txt
from orders_report import AmazonEUOrdersReport, AmazonCOMOrdersReport
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
EU_COUNTRIES_TXT = 'EU Countries.txt'
DPOST_REF_CHARLIMIT_PER_CELL = 28


class ParseOrders():
    '''Input: orders as list of dicts, parses orders, groups, forms output object;
    passes to OrdersReport class which creates report in xlsx format.
    Interacts with database client instance; main method:
    
    export_orders(testing=False) : groups orders by EU/ non-EU orders, with nesting based on currency.    
    when testing flag = True, export is suspended, but orders passed to class are still added to database
    
    Args:
    -orders : list - list of order dictionaries
    -amzn_channel : str - 'EU'/'COM' to differenciate different report
    -db_client:object - db client to iteract with during program runtime'''
    
    def __init__(self, all_orders : list, amzn_channel : str, db_client : object):
        self.all_orders = all_orders
        self.amzn_channel = amzn_channel
        self.db_client = db_client
        self.de_orders = []
        self.eu_orders = []
        self.non_eu_orders = []
    
    def _prepare_filepaths(self):
        '''creates cls variables of files abs paths to be created one dir above this script dir'''
        output_dir = get_output_dir()
        date_stamp = datetime.today().strftime("%Y.%m.%d %H.%M")
        self.report_path = os.path.join(output_dir, f'Amazon{self.amzn_channel} Report {date_stamp}.xlsx')
    
    def split_orders_by_region(self):
        '''based on amazon sales channel performs different regional sorting to class variables - lists. For Amazon EU:
        splits all_orders into three lists: EU (VAT (item-tax) > 0) and NON-EU (VAT = 0)
        for Amazon COM:
        splits all_orders into two lists: EU/NON-EU based on shipping-country compared to EU countries list'''
        self.eu_countries = self._get_EU_countries_list_from_file()
        for order in self.all_orders:
            try:
                if self.amzn_channel == 'COM':
                    # AMAZON COM regional sorting
                    if order['ship-country'] in self.eu_countries:
                        self.eu_orders.append(order)
                    else:
                        self.non_eu_orders.append(order)
                else:
                    # AMAZON EU regional sorting
                    # forcing UK orders in NON-EU group, independent from item-tax value (brexit update; also in __split_by_region in OrdersReport):
                    if order['ship-country'] == 'GB':
                        self.non_eu_orders.append(order)
                        continue

                    if order['ship-country'] == 'DE':
                        self.de_orders.append(order)
                        continue

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
    
    def _get_EU_countries_list_from_file(self):
        '''returns list of EU member countries from TXT file'''
        current_dir = get_output_dir(client_file=False)
        txt_abspath = os.path.join(current_dir, EU_COUNTRIES_TXT)
        logging.debug(f'Trying to access EU countries txt file: {txt_abspath}')
        return get_EU_countries_from_txt(txt_abspath)

    def exit_no_new_orders(self):
        '''terminate if all lists after sorting are empty'''
        if not self.eu_orders and not self.non_eu_orders and not self.de_orders:
            logging.info(f'No new orders found. Terminating, closing database connection, alerting VBA.')
            self.db_client.close_connection()
            print(VBA_NO_NEW_JOB)
            sys.exit()

    def prepare_export_obj(self):
        '''based on self.amzn_channel constructs a dict data object for OrdersReport class. Output format:
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
                # only for Amazon EU:
                de_orders: {
                            currency1: [order1, order2, order...],
                            currency2: [order1, order2, order...],
                            currency_n : [order1, order2, order...]
                                }                
                    }'''
        eu_currency_grouped = self.get_region_currency_based_dict(self.eu_orders)
        non_eu_currency_grouped = self.get_region_currency_based_dict(self.non_eu_orders)
        de_currency_grouped = self.get_region_currency_based_dict(self.de_orders)
        if self.amzn_channel == 'EU':
            self.export_obj = {'EU' : eu_currency_grouped, 'NON-EU' : non_eu_currency_grouped, 'DE' : de_currency_grouped}
        elif self.amzn_channel == 'COM':
            self.export_obj = {'EU' : eu_currency_grouped, 'NON-EU' : non_eu_currency_grouped}
        logging.debug(f'Returning export object with keys: {self.export_obj.keys()}')
        return self.export_obj

    @staticmethod
    def get_region_currency_based_dict(region_orders : list) -> dict:
        '''returns currency grouped dict.
        Example: {'EUR': [order1, order2...], 'USD':[order1, order2...], ...}'''
        currency_based_dict = defaultdict(list)
        for order in region_orders:
            currency_based_dict[order['currency']].append(order)
        return currency_based_dict

    def export_report(self):
        '''creates AmazonEUOrdersReport or AmazonCOMOrdersReport instance, and exports report in xlsx format'''
        try:
            if self.amzn_channel == 'EU':
                logging.info(f'Passing orders to create report with {AmazonEUOrdersReport.__name__} class')
                AmazonEUOrdersReport(self.export_obj).export(self.report_path)
            elif self.amzn_channel == 'COM':
                logging.info(f'Passing orders to create report with {AmazonCOMOrdersReport.__name__} class')
                AmazonCOMOrdersReport(self.export_obj, self.eu_countries).export(self.report_path)
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
        self.split_orders_by_region()
        self.exit_no_new_orders()
        self.prepare_export_obj()        
        if testing:
            logging.info(f'Running in testing {testing} environment. Change behaviour in export_orders method in ParseOrders class')
            print(f'Running in testing {testing} environment. Change behaviour in export_orders method in ParseOrders class')
            print('ENABLED REPORT EXPORT WHILE TESTING')            
            self.export_report()
            # self.push_orders_to_db()
            return
        self.export_report()
        self.push_orders_to_db()

if __name__ == "__main__":
    pass