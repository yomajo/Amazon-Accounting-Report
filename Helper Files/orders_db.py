from amzn_parser_utils import get_output_dir, file_to_binary, recreate_txt_file
from datetime import datetime
import sqlite3
import logging
import sys
import os


# GLOBAL VARIABLES
ORDERS_ARCHIVE_DAYS = 14
DATABASE_PATH = 'amzn_orders.db'
BACKUP_DB_BEFORE_NAME = 'amzn_orders_b4lrun.db'
BACKUP_DB_AFTER_NAME = 'amzn_orders_lrun.db'
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'


class OrdersDB:
    '''SQLite Database Management of Orders Flow. Takes list (list of dicts structure) of orders
    Two main methods:

    get_new_orders_only() - from passed orders to cls returns only ones, not yet in database.

    add_orders_to_db() - pushes new orders (return orders of get_new_orders_only() method) data
    selected data to database, performs backups before and after each run, periodic flushing of old entries'''
    
    def __init__(self, orders:list, txt_file_path:str):
        self.orders = orders
        self.txt_file_path = txt_file_path
        # database setup
        self.__get_db_paths()
        self.con = sqlite3.connect(self.db_path)
        self.con.execute("PRAGMA foreign_keys = 1;")
        self.con.execute('PRAGMA encoding="UTF-8";')
        self.__create_schema()
    
    def __get_db_paths(self):
        output_dir = get_output_dir(client_file=False)
        self.db_path = os.path.join(output_dir, DATABASE_PATH)
        self.db_backup_b4_path = os.path.join(output_dir, BACKUP_DB_BEFORE_NAME)
        self.db_backup_after_path = os.path.join(output_dir, BACKUP_DB_AFTER_NAME)

    def __create_schema(self):
        '''ensures 'program_runs' and 'orders' tables are in db'''
        try:
            with self.con:
                self.con.execute('''CREATE TABLE program_runs (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                    run_time TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                                                    weekday INTEGER,
                                                    fname TEXT DEFAULT 'No Fname Provided',
                                                    source_file BLOB);''')
        except sqlite3.OperationalError as e:
            logging.debug(f'program_runs table already created. Error: {e}')

        try:
            with self.con:
                self.con.execute('''CREATE TABLE orders (order_id TEXT PRIMARY KEY,
                                                purchase_date TEXT,
                                                payments_date TEXT,
                                                buyer_name TEXT NOT NULL,
                                                last_update TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                                                date_added TEXT NOT NULL,
                                                run INTEGER NOT NULL,
                                                FOREIGN KEY (run) REFERENCES program_runs (id) ON DELETE CASCADE);''')
        except sqlite3.OperationalError as e:
            logging.debug(f'orders table already created. Error: {e}')
        logging.debug('database tables are in place and ready to be used')

    def _get_order_ids_in_db(self) -> list:
        '''returns a list of order ids currently present in 'orders' database table'''
        try:
            with self.con:
                cur = self.con.cursor()
                cur.execute('''SELECT order_id FROM orders''')
                order_id_lst_in_db = [order_row[0] for order_row in cur.fetchall()]
                cur.close()
                logging.debug(f'Before inserting new orders, orders table contains {len(order_id_lst_in_db)} entries')
            return order_id_lst_in_db
        except sqlite3.OperationalError as e:
            logging.critical(f'Failed to retrieve order_ids as list from orders table. Syntax error: {e}')

    @staticmethod
    def get_today_weekday_int(date_arg=datetime.today()):
        '''returns integer for provided date (defaults to today). Monday - 1, ..., Sunday - 7'''
        return datetime.weekday(date_arg) + 1

    def _insert_new_run(self, weekday):
        '''Inserts new run (id, run_time, weekday, loaded filename and binary source) to program_runs table'''
        loaded_binary_file = file_to_binary(self.txt_file_path)
        try:
            with self.con:
                self.con.execute('''INSERT INTO program_runs (weekday, fname, source_file) VALUES (:weekday, :fname, :source_file)''',
                                {'weekday' : weekday,
                                'fname' : os.path.basename(self.txt_file_path),
                                'source_file' : loaded_binary_file})
                logging.debug(f'Added new run to program_runs table. Inserted with weekday: {weekday}')
        except Exception as e:
            logging.critical(f'Failed to insert new run to program_runs table. Error: {e}')

    def _get_current_run_id(self):
        '''return the most recent run_id by run_time column in db'''
        try:
            with self.con:
                cur = self.con.cursor()
                cur.execute('''SELECT id, run_time FROM program_runs ORDER BY run_time DESC LIMIT 1''')
                run_id, run_time = cur.fetchone()
                run_time_date = run_time.split(' ')[0]
                # Validaring the new run was made today (miliseconds before)
                assert run_time_date == datetime.today().strftime('%Y-%m-%d'), f'fetched run_time ({run_time_date}) date is not today'
                logging.info(f'Current program_runs id: {run_id}')
                return run_id
        except sqlite3.OperationalError as e:
            logging.error(f'Syntax error in query trying to fetch current run id. Error: {e}')

    def insert_multiple_orders(self, orders, run_id):
        '''adds all orders list members to 'orders' table in database. Assumes none of passed orders are in database'''
        date_added = datetime.today().strftime('%Y-%m-%d')
        for order_dict in orders:
            order_id = order_dict['order-id']
            purchase_date = order_dict['purchase-date']
            payments_date = order_dict['payments-date']
            buyer = order_dict['buyer-name']
            self.insert_new_order(order_id, purchase_date, payments_date, buyer, date_added, run_id)
        logging.info(f'{len(orders)} new orders were successfully added to database at run: {run_id}')

    def insert_new_order(self, order_id, purchase_date, payments_date, buyer_name, date_added, run_id):
        '''executes INSERT INTO orders with provided order args. Single order insert'''
        try:
            with self.con:
                self.con.execute('''INSERT INTO orders (order_id, purchase_date, payments_date, buyer_name, date_added, run)
                                                VALUES (:order_id, :purchase_date, :payments_date, :buyer_name, :date_added, :run)''',
                                                {'order_id':order_id, 'purchase_date':purchase_date, 'payments_date':payments_date,
                                                'buyer_name':buyer_name, 'date_added':date_added, 'run':run_id})
            logging.debug(f'Order {order_id} added to db successfully; run: {run_id} buyer: {buyer_name}')
        except sqlite3.OperationalError as e:
            logging.error(f'Order {order_id} insertion failed. Syntax error: {e}')
        except sqlite3.IntegrityError as e:
            logging.warning(f'Multi-item / multi-line order {order_id}. Already in database.')
        except Exception as e:
            logging.error(f'Unknown error while inserting order {order_id} data to orders table. Error: {e}')

    def __display_db_orders_table(self, order_by_last_update=False):
        '''debugging function. Prints out orders table to console and returns whole table as list of lists. Takes optional flag of timestamp sorting'''
        try:
            with self.con:
                cur = self.con.cursor()
                if order_by_last_update:   
                    cur.execute('''SELECT * FROM orders ORDER BY last_update DESC''')
                else:
                    cur.execute('''SELECT * FROM orders''')
                orders_table = cur.fetchall()
                for order_row in orders_table:
                    print(order_row)
                return orders_table
        except Exception as e:
            logging.error(f'Failed to retrieve data from orders table. Error {e}')

    def _flush_old_orders(self, archive_days=ORDERS_ARCHIVE_DAYS):
        '''cleans up database from orders added more than 'archive days' ago '''
        del_run_ids = self.__get_old_runs_ids(archive_days)
        try:
            with self.con:
                for run_id in del_run_ids:
                    self.con.execute('''DELETE FROM program_runs WHERE id = :run''', {'run':run_id})
            logging.info(f'Deleted old orders (cascade) from orders table where run_id = {del_run_ids}')
        except sqlite3.OperationalError as e:
            logging.error(f'Orders could not be deleted, passed run_ids: {del_run_ids}. Syntax error: {e}')
        except Exception as e:
            logging.error(f'Unknown error while deleting orders to orders table based on run_ids {del_run_ids}. Error: {e}')

    def __get_old_runs_ids(self, archive_days:int) -> list:
        '''returns list of run ids from program_runs table where runs were added more than 'archive_days' ago'''
        try:
            with self.con:
                cur = self.con.cursor()
                cur.execute('''SELECT id FROM program_runs WHERE
                            CAST(julianday('now', 'localtime') - julianday(run_time) AS INTEGER) >
                            :archive_days;''', {'archive_days':archive_days})
                old_run_ids = [run_row[0] for run_row in cur.fetchall()]
                cur.close()
            logging.debug(f'Identified old run ids: {old_run_ids}, added more than {archive_days} days ago')
            return old_run_ids
        except sqlite3.OperationalError as e:
            logging.error(f'Failed to retrieve ids from program_runs table. Syntax error: {e}')

    def _extract_file_to(self, output_dir:str, fname_in_db:str):
        '''recreates file from db table program runs where fname = fname_in_db, outputs file to output_dir'''
        output_abs_fpath = os.path.join(output_dir, fname_in_db)
        try:
            with self.con:
                cur = self.con.cursor()
                cur.execute('''SELECT source_file from program_runs WHERE fname = :fname_in_db''', {'fname_in_db':fname_in_db})
                sqlite_output = cur.fetchone()
            if sqlite_output == None:
                print(f'No database entry in program_runs table where fname = {fname_in_db}')
                self.close_connection()
                sys.exit()
            fetched_f_bin_data = sqlite_output[0]
            recreate_txt_file(output_abs_fpath, fetched_f_bin_data)
            print(f'Successfully recreated file {os.path.basename(output_abs_fpath)} from db to {output_abs_fpath}')
        except Exception as e:
            print(f'Unknown error encountered while retrieving file {fname_in_db} from db. Err: {e}')
        finally:
            print('Closing connection')
            self.close_connection()


    def _backup_db(self, backup_db_path):
        '''if everything is ok, backups could be performed weekly adding conditional:
        if self.get_today_weekday_int() == 5:'''
        back_con = sqlite3.connect(backup_db_path)
        with back_con:
            self.con.backup(back_con, pages=0, name='main')
        back_con.close()
        logging.info(f"New database backup {os.path.basename(backup_db_path)} created on: "
                    f"{datetime.today().strftime('%Y-%m-%d %H:%M')} location: {backup_db_path}")
    
    def close_connection(self):
        self.con.close()
        logging.info(f'Connection to DB in session with file {os.path.basename(self.txt_file_path)} closed')


    def get_new_orders_only(self):
        '''From passed orders to cls, returns only ones NOT YET in database'''
        orders_in_db = self._get_order_ids_in_db()
        self.new_orders = [order_data for order_data in self.orders if order_data['order-id'] not in orders_in_db]
        logging.info(f'Returning {len(self.new_orders)}/{len(self.orders)} new/loaded orders for further processing')
        logging.debug(f'Database currently holds {len(orders_in_db)} order records')
        return self.new_orders

    def add_orders_to_db(self):
        '''adds all cls orders to db, flushes old records, performs backups before and after changes to db,
        returns number of orders added to db'''
        try:
            self._backup_db(self.db_backup_b4_path)
            logging.info(f'Created backup {os.path.basename(self.db_backup_b4_path)} before adding orders')
            # Adding new orders:
            self._insert_new_run(self.get_today_weekday_int())
            new_run_id = self._get_current_run_id()
            self.insert_multiple_orders(self.new_orders, new_run_id)
            # House keeping
            self._flush_old_orders(ORDERS_ARCHIVE_DAYS)
            self._backup_db(self.db_backup_after_path)
            logging.info(f'Created backup {os.path.basename(self.db_backup_after_path)} after adding orders')
            logging.debug('add_orders_to_db finished successfully. Both backups created.')
        except Exception as e:
            logging.critical(f'Unknown error when inserting new orders. Error: {e}. Alerting VBA side about errors')
            print(VBA_ERROR_ALERT)
        finally:
            self.close_connection()
        return len(self.new_orders)


if __name__ == "__main__":
    pass