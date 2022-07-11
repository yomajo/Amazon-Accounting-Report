# ---------------- NEW AT TOP, UNREVIEWED -------------------------
import datetime
import logging
import os
import shutil
from sqlalchemy import create_engine, Column, String, Integer
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy.sql.sqltypes import TIMESTAMP
from sqlalchemy.sql.schema import ForeignKey
from sqlalchemy.exc import IntegrityError
from utils import get_output_dir, create_src_file_backup, delete_file


# GLOBAL VARIABLES
ORDERS_ARCHIVE_DAYS = 120
DATABASE_NAME = 'inventory.db'
BACKUP_DB_BEFORE_NAME = 'inventory_b4lrun.db'
BACKUP_DB_AFTER_NAME = 'inventory_lrun.db'
VBA_ERROR_ALERT = 'ERROR_CALL_DADDY'

Base = declarative_base()


class ProgramRun(Base):
    '''database table model representing unique program run'''
    __tablename__ = 'program_run'

    def __init__(self, fpath:str, sales_channel, timestamp=datetime.datetime.now(), **kwargs):
        super(ProgramRun, self).__init__(**kwargs)
        self.fpath = fpath
        self.sales_channel = sales_channel
        self.timestamp = timestamp

    id = Column(Integer, primary_key=True, nullable=False)
    fpath = Column(String, nullable=False)
    sales_channel = Column(String, nullable=False)      # Amazon /Amazon Warehouse /Etsy
    timestamp = Column(TIMESTAMP(timezone=False), default=datetime.datetime.now())
    orders = relationship('Order', cascade='all, delete', cascade_backrefs=True,
                passive_deletes=False, passive_updates=False, backref='run_obj')

    def __repr__(self) -> str:
        return f'<ProgramRun id: {self.id}, sales_channel: {self.sales_channel}, timestamp: {self.timestamp}, fpath: {self.fpath}>'
    

class Order(Base):
    '''database table model representing Order
    
    NOTE: unique primary key is:
        order['order-item-id'] for Amazon;
        order['Shipment Item ID'] for Amazon Warehouse;
        order['Order ID'] for Etsy;

        order_id_secondary:
        order['order-id'] for Amazon;
        order['Amazon Order Id'] for Amazon Warehouse;
        null for Etsy'''
    __tablename__ = 'order'

    def __init__(self, order_id, purchase_date, buyer_name, run, **kwargs):
        super(Order, self).__init__(**kwargs)
        self.order_id = order_id
        self.purchase_date = purchase_date
        self.buyer_name = buyer_name
        self.run = run

    order_id = Column(String, primary_key=True, nullable=False)
    order_id_secondary = Column(String)
    purchase_date = Column(String)
    buyer_name = Column(String)
    run = Column(Integer, ForeignKey('program_run.id', ondelete='CASCADE', onupdate='CASCADE'), nullable=False)

    def __repr__(self) -> str:
        return f'<Order order_id: {self.order_id}, added on run: {self.run}>'


class SQLAlchemyOrdersDB:
    '''Orders Database management. Two main methods:

    get_new_orders_only() - from passed orders to cls returns only ones, not yet in database.
    Expected to be called outside of this cls to get self.new_orders var.

    add_orders_to_db() - pushes new orders (returned list from get_new_orders_only() method)
    selected data to database, performs backups before and after each run, periodic flushing of old entries 
    
    IMPORTANT NOTE: Amazon has unique order-item-id's (same order-id for different items in buyer's cart).
    Order model saves order['order-item-id'] for Amazon orders and for Etsy: order['Order ID']
    
    Arguments:

    orders - list of dict / OrderedDict's

    source_file_path - abs path to source file for orders (Amazon / Etsy)

    sales_channel - str identifier for db entry, backup file naming. Expected value: ['Amazon', 'Amazon Warehouse', 'Etsy']

    proxy_keys - dict mapper of internal (based on amazon) order keys vs external sales_channel keys 

    testing - optional flag for testing (suspending backup, save add source_file_path to program_run table instead)
    '''

    def __init__(self, orders:list, source_file_path:str, sales_channel:str, proxy_keys:dict, testing=False):
        self.orders = orders
        self.source_file_path = source_file_path
        self.sales_channel = sales_channel
        self.proxy_keys = proxy_keys
        self.testing = testing
        self.__setup_db()
        self._backup_db(self.db_backup_b4_path)
        self.session = self.get_session()

    def __setup_db(self):
        self.__get_db_paths()
        if not os.path.exists(self.db_path):
            self.__get_engine()
            Base.metadata.create_all(bind=self.engine)
            logging.info(f'Database has been created at {self.db_path}')

    def __get_db_paths(self):
        output_dir = get_output_dir(client_file=False)
        self.db_path = os.path.join(output_dir, DATABASE_NAME)
        self.db_backup_b4_path = os.path.join(output_dir, BACKUP_DB_BEFORE_NAME)
        self.db_backup_after_path = os.path.join(output_dir, BACKUP_DB_AFTER_NAME)

    def __get_engine(self):
        engine_path = f'sqlite:///{self.db_path}'
        self.engine = create_engine(engine_path, echo=False)
    
    def get_session(self):
        '''returns database session object to work outside the scope of class. For example querying'''
        self.__get_engine()
        Session = sessionmaker(bind=self.engine)
        return Session()

    def add_orders_to_db(self):
        '''filters passed orders to cls to only those, whose order_id
        (db table unique constraint) is not present in db yet adds them to db
        assumes get_new_orders_only was called outside of this cls before to get self.new_orders'''
        try:
            if self.new_orders:
                self._add_new_orders_to_db(self.new_orders)
                self.flush_old_records()
                self._backup_db(self.db_backup_after_path)
            logging.debug(f'{len(self.new_orders)} (order count) new orders added, flushing old records complete, backup after created at: {self.db_backup_after_path}')
            return len(self.new_orders)
        except Exception as e:
            logging.critical(f'Unexpected err {e} trying to add orders to db. Alerting VBA, terminating program immediately via exit().')
            print(VBA_ERROR_ALERT)
            exit()

    def _add_new_orders_to_db(self, new_orders:list):
        '''create new entry in program_runs table, add new orders'''
        self.new_run = self._add_new_run()
        self.added_to_db_counter = 0
        for order in new_orders:
            self._add_single_order(order)
        logging.debug(f'{self.added_to_db_counter} new orders added to db (actual counter of commits)')
            

    def _add_single_order(self, order_dict:dict):
        '''adds single order to database (via session.add(new_order))'''
        try:
            new_order = Order(order_id = order_dict[self.proxy_keys['order-id']],
                            purchase_date = order_dict[self.proxy_keys['purchase-date']],
                            buyer_name = order_dict[self.proxy_keys['buyer-name']],
                            run = self.new_run.id)
            if self.new_run.sales_channel != 'Etsy':
                # Additionally add original order-id (may have duplicates for multiple items in shopping cart) for AmazonCOM, AmazonEU
                # Both Amazon and Amazon Warehouse have 'secondary-order-id' secondary key
                new_order.order_id_secondary = order_dict[self.proxy_keys['secondary-order-id']]
            
            self.session.add(new_order)
            self.session.commit()
            self.added_to_db_counter += 1
        except IntegrityError as e:
            logging.warning(f'Order from channel: {self.sales_channel} w/ proxy order-id: {order_dict[self.proxy_keys["order-id"]]} \
                already in database. Integrity error {e}. Skipping addition of said order, rolling back db session')
            self.session.rollback()

    def _add_new_run(self) -> object:
        '''adds new row in program_run table, returns new run object (attributes: id, sales_channel, fpath, timestamp),
        creates source file backup, saves its path. On testing - save original file path'''        
        backup_path = self.source_file_path if self.testing else create_src_file_backup(self.source_file_path, self.sales_channel)
        logging.debug(f'This is backup path being saved to program_run fpath column: {backup_path}')
        new_run = ProgramRun(fpath=backup_path, sales_channel=self.sales_channel)
        self.session.add(new_run)
        self.session.commit()
        logging.debug(f'Added new run: {new_run}, created backup')
        return new_run

    def get_new_orders_only(self) -> list:
        '''From passed orders to cls, returns only orders NOT YET in database.
        Called from main.py to filter old, parsed orders'''
        orders_in_db = self._get_channel_order_ids_in_db()
        self.new_orders = [order_data for order_data in self.orders if order_data[self.proxy_keys['order-id']] not in orders_in_db]
        logging.info(f'Returning {len(self.new_orders)}/{len(self.orders)} new/loaded orders for further processing')
        return self.new_orders

    def _get_channel_order_ids_in_db(self) -> list:
        '''returns a list of order ids currently present in 'orders' database table for current run self.sales_channel'''
        db_orders_of_sales_channel = self.session.query(Order).join(ProgramRun).filter(ProgramRun.sales_channel==self.sales_channel).all()
        # Unlikely conflict: Etsy / Amazon EU having same order-(item-)id as AmazonCOM or similar permutations between sales channels and id's
        order_id_lst_in_db = [order_obj.order_id for order_obj in db_orders_of_sales_channel]
        logging.debug(f'Before inserting new orders, orders table contains {len(order_id_lst_in_db)} entries associated with {self.sales_channel} channel')
        return order_id_lst_in_db

    def flush_old_records(self):
        '''deletes old runs, associated backup files and orders (deleting runs delete cascade associated orders)'''
        old_runs = self._get_old_runs()
        try:
            for run in old_runs:
                orders_in_run = self.session.query(Order).filter_by(run_obj=run).all()
                logging.info(f'Deleting {len(orders_in_run)} orders associated with old {run} and backup file: {run.fpath}')
                delete_file(run.fpath)   
                self.session.delete(run)
            self.session.commit()
        except Exception as e:
            logging.warning(f'Unexpected err while flushing old records from db inside flush_old_records. Err: {e}. Last recorded run {run}')

    def _get_old_runs(self):
        '''returns runs that were added ORDERS_ARCHIVE_DAYS (global var) or more days ago'''
        delete_before_this_timestamp = datetime.datetime.now() - datetime.timedelta(days=ORDERS_ARCHIVE_DAYS)        
        runs = self.session.query(ProgramRun).filter(ProgramRun.timestamp < delete_before_this_timestamp).all()
        return runs

    def _backup_db(self, backup_db_path):
        '''creates database backup file at backup_db_path in production (testing = False)'''
        if self.testing:
            logging.debug(f'Backup for {os.path.basename(backup_db_path)} suspended due to testing: {self.testing}')
            return
        try:
            shutil.copy(src=self.db_path, dst=backup_db_path)
            logging.info(f"New database backup {os.path.basename(backup_db_path)} created on: "
                        f"{datetime.datetime.today().strftime('%Y-%m-%d %H:%M')} location: {backup_db_path}")
        except Exception as e:
            logging.warning(f'Failed to create database backup for {os.path.basename(backup_db_path)}. Err: {e}')


if __name__ == "__main__":
    pass


# ------------------------------------------------------------
# ------------------------ OLD BELOW -------------------------
# ------------------------------------------------------------
import sqlite3
import logging
import sys
import os
from datetime import datetime
from utils import get_output_dir, file_to_binary, recreate_txt_file
from constants import VBA_ERROR_ALERT


# GLOBAL VARIABLES
ORDERS_ARCHIVE_DAYS = 90
DATABASE_PATH = 'amzn_accounting.db'
BACKUP_DB_BEFORE_NAME = 'amzn_accounting_b4lrun.db'
BACKUP_DB_AFTER_NAME = 'amzn_accounting_lrun.db'


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
                self.con.execute('''CREATE TABLE orders (order_id TEXT,
                                                order_item_id TEXT,
                                                purchase_date TEXT,
                                                payments_date TEXT,
                                                buyer_name TEXT NOT NULL,
                                                quantity INTEGER,
                                                currency TEXT,
                                                item_price DECIMAL,
                                                item_tax DECIMAL,
                                                shipping_price DECIMAL,
                                                shipping_tax DECIMAL,
                                                last_update TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                                                date_added TEXT NOT NULL,
                                                run INTEGER NOT NULL,
                                                PRIMARY KEY (order_id, order_item_id),
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
        for order in orders:
            self.insert_new_order(order, date_added, run_id)
        logging.info(f'{len(orders)} new orders were successfully added to database at run: {run_id}')

    def insert_new_order(self, order : dict, date_added : str, run_id : str):
    # def insert_new_order(self, order_id, purchase_date, payments_date, buyer_name, date_added, run_id):
        '''executes INSERT INTO 'orders' table with provided run_id, data for insert from order (dict). Single order insert'''
        order_id = order['order-id']
        buyer_name = order['buyer-name']
        try:
            with self.con:
                self.con.execute('''INSERT INTO orders (order_id, order_item_id, purchase_date, payments_date, buyer_name,
                                quantity, currency, item_price, item_tax, shipping_price, shipping_tax, date_added, run)
                                
                                VALUES (:order_id, :order_item_id, :purchase_date, :payments_date, :buyer_name, :quantity,
                                :currency, :item_price, :item_tax, :shipping_price, :shipping_tax,:date_added, :run)''',

                                {'order_id':order_id, 'order_item_id':order['order-item-id'], 'purchase_date':order['purchase-date'],
                                'payments_date':order['payments-date'], 'buyer_name':order['buyer-name'], 'quantity':order['quantity-purchased'],
                                'currency':order['currency'], 'item_price':order['item-price'], 'item_tax':order['item-tax'],
                                'shipping_price':order['shipping-price'], 'shipping_tax':order['shipping-tax'], 'date_added':date_added, 'run':run_id})
            logging.debug(f'Order {order_id} added to db successfully; run: {run_id} buyer: {buyer_name}')
        except sqlite3.OperationalError as e:
            logging.error(f'Order {order_id} insertion failed. Syntax error: {e}')
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