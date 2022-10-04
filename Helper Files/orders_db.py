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
from accounting_utils import get_output_dir, create_src_file_backup, delete_file
from constants import VBA_ERROR_ALERT


# GLOBAL VARIABLES
ORDERS_ARCHIVE_DAYS = 120
DATABASE_PATH = 'amzn_accounting.db'
BACKUP_DB_BEFORE_NAME = 'amzn_accounting_b4lrun.db'
BACKUP_DB_AFTER_NAME = 'amzn_accounting_lrun.db'


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
    sales_channel = Column(String, nullable=False)      # AmazonEU / AmazonCOM / Amazon Warehouse
    timestamp = Column(TIMESTAMP(timezone=False), default=datetime.datetime.now())
    orders = relationship('Order', cascade='all, delete', cascade_backrefs=True,
                passive_deletes=False, passive_updates=False, backref='run_obj')

    def __repr__(self) -> str:
        return f'<ProgramRun id: {self.id}, sales_channel: {self.sales_channel}, timestamp: {self.timestamp}, fpath: {self.fpath}>'
    

class Order(Base):
    '''database table model representing Order
    
    NOTE: unique primary key is:
        order['order-item-id'] for Amazon (AmazonEU / AmazonCOM);
        order['Shipment Item ID'] for Amazon Warehouse;

        order_id_secondary:
        order['order-id'] for Amazon;
        order['Amazon Order Id'] for Amazon Warehouse;
    '''
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
    Order model saves order['order-item-id'] for Amazon orders
    
    Arguments:

    orders - list of dict / OrderedDict's

    source_file_path - abs path to source file for orders AmazonCOM / AmazonEU / Amazon Warehouse)

    sales_channel - str identifier for db entry, backup file naming. Expected value: ['AmazonCOM', 'AmazonEU', 'Amazon Warehouse']

    proxy_keys - dict mapper of internal (based on amazon) order keys vs external sales_channel keys 

    testing - optional flag for testing (suspending backup, save add source_file_path to program_run table instead)
    '''

    def __init__(self, orders: list, source_file_path: str, sales_channel: str, proxy_keys: dict, testing: bool=False):
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
        self.db_path = os.path.join(output_dir, DATABASE_PATH)
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
            # Leaving, in case Etsy gets integrated at some point in the future
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
        Called from main_accounting.py to filter old, parsed orders'''
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
            self.session.rollback()

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
    
    def close_connection(self):
        '''closes db session'''
        self.session.close()


if __name__ == "__main__":
    pass