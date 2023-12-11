"""
This module contains classes for creating tables of the database
and database related functions like create table, create database,
check empty values in existing tables.

"""

import os
import sys
from datetime import datetime
from sqlalchemy import ForeignKey, Column, String, Integer, Date, DECIMAL, create_engine, update, and_
from sqlalchemy.schema import UniqueConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import pandas as pd

# Base class for SQL objects to extent
Base = declarative_base()

border_line = '🂱 ♠ 🂱 ♥️ 🂱 ♣️ 🂱 ♦️ 🂱 ♠ 🂱 ♥️ 🂱 ♣️ 🂱 ♦️ 🂱 ♠ 🂱 ♥️ 🂱 '
border_line += '🂱 ♠ 🂱 ♥️ 🂱 ♣️ 🂱 ♦️ 🂱 ♠ 🂱 ♥️ 🂱 ♣️ 🂱 ♦️ 🂱 ♠ 🂱 ♥️ 🂱 '

# Create classes for the tables as objects
class Companies(Base):

    """
    This class allows to create and manipulate table named 'companies' in the SQL database
    """

    __tablename__ = 'companies'

    company_id = Column('company_id', Integer, primary_key=True, nullable=False)
    company_name = Column('company_name', String, nullable=False)
    company_address = Column('company_address', String)
    company_branch = Column('company_branch', String)

    __table_args__ = (UniqueConstraint('company_id', 'company_name'),)

    def __init__(self, dataframe_table, row_index):
        self.company_name = dataframe_table['company_name'][row_index]
        self.company_address = dataframe_table['company_address'][row_index]
        self.company_branch = dataframe_table['company_branch'][row_index]

    def __repr__(self):
        return (f"COMPANIES\n"
                f"ID: {self.company_id},\n"
                f"Name: {self.company_name},\n"
                f"Branch: {self.company_branch},\n"
                f"Address: {self.company_address}\n")

class CurrencyUsd(Base):

    """
    This class allows to create and manipulate table named 'currency_usd' in the SQL database
    """

    __tablename__ = 'currency_usd'

    currency_usd_id = Column('currency_usd_id', Integer, primary_key=True, nullable=False)
    currency_type = Column('currency_type', String, nullable=False)
    currency_date = Column('currency_date', String, nullable=False, unique=True)
    currency_rate = Column('currency_rate', DECIMAL, nullable=False)

    __table_args__ = (UniqueConstraint('currency_usd_id', 'currency_date'),)

    def __init__(self, dataframe_table, row_index):
        self.currency_date = dataframe_table['currency_date'][row_index]
        self.currency_rate = dataframe_table['currency_rate'][row_index]
        self.currency_type = 'usd'

    def __repr__(self):
        return (f"CURRENCY USD\n"
                f"ID: {self.currency_usd_id},\n"
                f"Type: {self.currency_type},\n"
                f"Date: {self.currency_date},\n"
                f"Rate: {self.currency_rate}\n")

class CurrencyEur(Base):

    """
    This class allows to create and manipulate table named 'currency_eur' in the SQL database
    """

    __tablename__ = 'currency_eur'

    currency_eur_id = Column('currency_eur_id', Integer, primary_key=True, nullable=False)
    currency_type = Column('currency_type', String, nullable=False)
    currency_date = Column('currency_date', String, nullable=False, unique=True)
    currency_rate = Column('currency_rate', DECIMAL, nullable=False)

    __table_args__ = (UniqueConstraint('currency_eur_id', 'currency_date'),)

    def __init__(self, dataframe_table, row_index):
        self.currency_date = dataframe_table['currency_date'][row_index]
        self.currency_rate = dataframe_table['currency_rate'][row_index]
        self.currency_type = 'eur'

    def __repr__(self):
        return (f"CURRENCY EUR\n"
                f"ID: {self.currency_eur_id},\n"
                f"Type: {self.currency_type},\n"
                f"Date: {self.currency_date},\n"
                f"Rate: {self.currency_rate}\n")

class Customers(Base):

    """
    This class allows to create and manipulate table named 'customers' in the SQL database
    """

    __tablename__ = 'customers'

    customer_id = Column('customer_id', Integer, primary_key=True, nullable=False)
    customer_name = Column('customer_name', String, nullable=False)
    customer_address = Column('customer_address', String)
    customer_phone = Column('customer_phone', String)
    customer_email = Column('customer_email', String)
    customer_cons_kg = Column('customer_cons_kg', DECIMAL)
    customer_eq_kg = Column('customer_eq_kg', DECIMAL)
    customer_ken_kg = Column('customer_ken_kg', DECIMAL)
    customer_col_kg = Column('customer_col_kg', DECIMAL)
    customer_isr_kg = Column('customer_isr_kg', DECIMAL)
    customer_isr_pallet = Column('customer_isr_pallet', Integer)
    customer_holl_pallet = Column('customer_holl_pallet', Integer)
    customer_preecool_kg = Column('customer_preecool_kg', DECIMAL)
    customer_preecool_awb = Column('customer_preecool_awb', Integer)
    customer_flight_kg = Column('customer_flight_kg', DECIMAL)
    customer_troll = Column('customer_troll', Integer)
    customer_bulb_pallet = Column('customer_bulb_pallet', Integer)
    customer_rus_eq_full = Column('customer_rus_eq_full', Integer)
    customer_rus_else_full = Column('customer_rus_else_full', Integer)
    customer_rus_big_box = Column('customer_rus_big_box', Integer)
    customer_rus_small_box = Column('customer_rus_small_box', Integer)
    customer_dollar_trans_rate = Column('customer_dollar_trans_rate', Integer)
    customer_euro_trans_rate = Column('customer_euro_trans_rate', Integer)
    customer_dollar_flow_rate = Column('customer_dollar_flow_rate', Integer)
    customer_euro_flow_rate = Column('customer_euro_flow_rate', Integer)
    customer_flow_markup = Column('customer_flow_markup', DECIMAL)
    customer_trans_markup = Column('customer_trans_markup', DECIMAL)

    __table_args__ = (UniqueConstraint('customer_id', 'customer_name',
                                       'customer_phone', 'customer_email'),)

    def __init__(self, dataframe_table, row_index):
        self.customer_name = dataframe_table['customer_name'][row_index]
        self.customer_address = dataframe_table['customer_address'][row_index]
        self.customer_phone = dataframe_table['customer_phone'][row_index]
        self.customer_email = dataframe_table['customer_email'][row_index]
        self.customer_cons_kg = dataframe_table['customer_cons_kg_price'][row_index]
        self.customer_eq_kg = dataframe_table['customer_eq_kg_price'][row_index]
        self.customer_ken_kg = dataframe_table['customer_ken_kg_price'][row_index]
        self.customer_col_kg = dataframe_table['customer_col_kg_price'][row_index]
        self.customer_isr_kg = dataframe_table['customer_isr_kg_price'][row_index]
        self.customer_isr_pallet = dataframe_table['customer_isr_pallet_price'][row_index]
        self.customer_holl_pallet = dataframe_table['customer_holl_pallet_price'][row_index]
        self.customer_preecool_kg = dataframe_table['customer_preecool_kg_price'][row_index]
        self.customer_preecool_awb = dataframe_table['customer_preecool_awb_price'][row_index]
        self.customer_flight_kg = dataframe_table['customer_flight_kg_price'][row_index]
        self.customer_troll = dataframe_table['customer_troll_price'][row_index]
        self.customer_bulb_pallet = dataframe_table['customer_bulb_pallet_price'][row_index]
        self.customer_rus_eq_full = dataframe_table['customer_rus_eq_full_price'][row_index]
        self.customer_rus_else_full = dataframe_table['customer_rus_else_full_price'][row_index]
        self.customer_rus_big_box = dataframe_table['customer_rus_big_box_price'][row_index]
        self.customer_rus_small_box = dataframe_table['customer_rus_small_box_price'][row_index]
        self.customer_dollar_trans_rate = dataframe_table['customer_dollar_trans_rate'][row_index]
        self.customer_euro_trans_rate = dataframe_table['customer_euro_trans_rate'][row_index]
        self.customer_dollar_flow_rate = dataframe_table['customer_dollar_flow_rate'][row_index]
        self.customer_euro_flow_rate = dataframe_table['customer_euro_flow_rate'][row_index]
        self.customer_flow_markup = dataframe_table['customer_flow_markup'][row_index]
        self.customer_trans_markup = dataframe_table['customer_trans_markup'][row_index]

    def __repr__(self):
        return (f"CUSTOMERS\n"
                f"ID: {self.customer_id},\n"
                f"Name: {self.customer_name},\n"
                f"Address: {self.customer_address},\n"
                f"Phone: {self.customer_phone},\n"
                f"Email: {self.customer_email},\n"
                f"Consolidation EC/kg: {self.customer_cons_kg},\n"
                f"Import Car EC/kg: {self.customer_eq_kg},\n"
                f"Import Car KEN/kg: {self.customer_ken_kg},\n"
                f"Import Car COL/kg: {self.customer_col_kg},\n"
                f"Import Car ISR/kg: {self.customer_isr_kg},\n"
                f"Import Car ISR/palett: {self.customer_isr_pallet},\n"
                f"Import Car HOL/palett: {self.customer_holl_pallet},\n"
                f"Pre-cooling/kg: {self.customer_preecool_kg},\n"
                f"Pre-cooling/awb: {self.customer_preecool_awb},\n"
                f"Flight EC/kg: {self.customer_flight_kg},\n"
                f"Import Car troll/piece: {self.customer_troll},\n"
                f"Import Car bulb/palett: {self.customer_bulb_pallet},\n"
                f"Local Car EC/full: {self.customer_rus_eq_full},\n"
                f"Local Car NON-EC/full: {self.customer_rus_else_full},\n"
                f"Local Car big HOL/piece: {self.customer_rus_big_box},\n"
                f"Local Car small HOL/piece: {self.customer_rus_small_box},\n"
                f"Currency for transport/USD: {self.customer_dollar_trans_rate},\n"
                f"Currency for transport/EUR: {self.customer_euro_trans_rate},\n"
                f"Currency for flowers/USD: {self.customer_dollar_flow_rate},\n"
                f"Currency for flowers/EUR: {self.customer_euro_flow_rate},\n"
                f"Markup for flowers/%: {self.customer_flow_markup},\n"
                f"Markup for transport/%: {self.customer_trans_markup}\n")

class Markings(Base):

    """
    This class allows to create and manipulate table named 'markings' in the SQL database
    """

    __tablename__ = 'markings'

    marking_id = Column('marking_id', Integer, primary_key=True, nullable=False)
    marking_customer = Column('marking_customer', String, nullable=False)
    marking_customer_address = Column('marking_customer_address', String,)
    marking_name = Column('marking_name', String, nullable=False)
    customer_id = Column(Integer, ForeignKey('customers.customer_id'))

    __table_args__ = (UniqueConstraint('marking_id', 'marking_name'),)

    def __init__(self, dataframe_table, row_index):
        self.marking_customer = dataframe_table['marking_customer'][row_index]
        self.marking_customer_address = dataframe_table['marking_customer_address'][row_index]
        self.marking_name = dataframe_table['marking_name'][row_index]

    def __repr__(self):
        return (f"MARKINGS\n"
                f"ID: {self.marking_id},\n"
                f"Customer: {self.marking_customer},\n"
                f"Address: {self.marking_customer_address},\n"
                f"Marking: {self.marking_name},\n"
                f"Customer's ID: {self.customer_id}\n")

class Suppliers(Base):

    """
    This class allows to create and manipulate table named 'suppliers' in the SQL database
    """

    __tablename__ = 'suppliers'

    supplier_id = Column('supplier_id', Integer, primary_key=True, nullable=False)
    supplier_name = Column('supplier_name', String, nullable=False)
    supplier_country = Column('supplier_country', String)

    __table_args__ = (UniqueConstraint('supplier_id', 'supplier_name'),)

    def __init__(self, dataframe_table, row_index):
        self.supplier_name = dataframe_table['supplier_name'][row_index]
        self.supplier_country = dataframe_table['supplier_country'][row_index]

    def __repr__(self):
        return (f"SUPPLIERS\n"
                f"ID: {self.supplier_id},\n"
                f"Name: {self.supplier_name},\n"
                f"Country: {self.supplier_country}\n")

class BoxType(Base):

    """
    This class allows to create and manipulate table named 'box_type' in the SQL database
    """

    __tablename__ = 'box_type'

    box_type_id = Column('box_type_id', Integer, primary_key=True, nullable=False)
    box_name = Column('box_name', String, nullable=False)
    box_per_pallet = Column('box_per_pallet', Integer, nullable=False)
    box_type_accountable = Column('box_type_accountable', String, nullable=False)

    __table_args__ = (UniqueConstraint('box_type_id', 'box_name'),)

    def __init__(self, dataframe_table, row_index):
        self.box_name = dataframe_table['box_name'][row_index]
        self.box_per_pallet = dataframe_table['box_per_pallet'][row_index]
        self.box_type_accountable = dataframe_table['box_type_accountable'][row_index]

    def __repr__(self):
        return (f"BOX TYPE\n"
                f"ID: {self.box_type_id},\n"
                f"Name: {self.box_name},\n"
                f"Boxes per palett: {self.box_per_pallet},\n"
                f"Returnal type: {self.box_type_accountable}\n")

class FlowerType(Base):

    """
    This class allows to create and manipulate table named 'flower_type' in the SQL database
    """

    __tablename__ = 'flower_type'

    flower_id = Column('flower_id', Integer, primary_key=True, nullable=False)
    flower_name = Column('flower_name', String, nullable=False)
    flower_type = Column('flower_type', String, nullable=False)
    flower_plantation = Column('flower_plantation', String)

    __table_args__ = (UniqueConstraint('flower_id', 'flower_name'),)

    def __init__(self, dataframe_table, row_index):
        self.flower_name = dataframe_table['flower_name'][row_index]
        self.flower_type = dataframe_table['flower_type'][row_index]
        self.flower_plantation = dataframe_table['flower_plantation'][row_index]

    def __repr__(self):
        return (f"FLOWER TYPE\n"
                f"ID: {self.flower_id},\n"
                f"Name: {self.flower_name},\n"
                f"Type: {self.flower_type},\n"
                f"Plantation: {self.flower_plantation}\n")

class Managers(Base):

    """
    This class allows to create and manipulate table named 'managers' in the SQL database
    """

    __tablename__ = 'managers'

    manager_id = Column('manager_id', Integer, primary_key=True, nullable=False)
    manager_name = Column('manager_name', String, nullable=False)
    manager_birth = Column('manager_birth', String)
    manager_salary = Column('manager_salary', Integer, nullable=False)
    manager_sales_perc = Column('manager_sales_perc', Integer, nullable=False)
    company_id = Column(Integer, ForeignKey('companies.company_id'))

    __table_args__ = (UniqueConstraint('manager_id', 'manager_name'),)

    def __init__(self, dataframe_table, row_index):
        self.manager_name = dataframe_table['manager_name'][row_index]
        self.manager_birth = dataframe_table['manager_birth'][row_index]
        self.manager_salary = dataframe_table['manager_salary'][row_index]
        self.manager_sales_perc = dataframe_table['manager_sales_perc'][row_index]
        self.company_id = dataframe_table['company_id'][row_index]

    def __repr__(self):
        return (f"MANAGERS\n"
                f"ID: {self.manager_id},\n"
                f"Name: {self.manager_name},\n"
                f"Date of birth: {self.manager_birth},\n"
                f"Salary: {self.manager_salary},\n"
                f"Sale's percent: {self.manager_sales_perc},\n"
                f"Company's ID: {self.company_id}\n")

class Cars(Base):

    """
    This class allows to create and manipulate table named 'cars' in the SQL database
    """

    __tablename__ = 'cars'

    car_id = Column('car_id', Integer, primary_key=True, nullable=False)
    car_brand = Column('car_brand', String, nullable=False)
    car_plate = Column('car_plate', String, nullable=False)
    car_year = Column('car_year', Integer, nullable=False)
    car_mileage = Column('car_mileage', Integer, nullable=False)
    car_capacity = Column('car_capacity', Integer, nullable=False)
    company_id = Column(Integer, ForeignKey('companies.company_id'), nullable=False)

    __table_args__ = (UniqueConstraint('car_id', 'car_plate'),)

    def __init__(self, dataframe_table, row_index):
        self.car_brand = dataframe_table['car_brand'][row_index]
        self.car_plate = dataframe_table['car_plate'][row_index]
        self.car_year = dataframe_table['car_year'][row_index]
        self.car_mileage = dataframe_table['car_mileage'][row_index]
        self.car_capacity = dataframe_table['car_capacity'][row_index]
        self.company_id = dataframe_table['company_id'][row_index]

    def __repr__(self):
        return (f"CARS\n"
                f"ID: {self.car_id},\n"
                f"Brand: {self.car_brand},\n"
                f"Plate: {self.car_plate},\n"
                f"Year: {self.car_year},\n"
                f"Mileage: {self.car_mileage},\n"
                f"Capacity: {self.car_capacity},\n"
                f"Company's ID: {self.company_id}\n")

class Drivers(Base):

    """
    This class allows to create and manipulate table named 'drivers' in the SQL database
    """

    __tablename__ = 'drivers'

    driver_id = Column('driver_id', Integer, primary_key=True, nullable=False)
    driver_name = Column('driver_name', String, nullable=False)
    driver_phone = Column('driver_phone', Integer)
    driver_email = Column('driver_email', String)
    driver_passport = Column('driver_passport', String)
    driver_address = Column('driver_address', String)
    driver_license = Column('driver_license', String)
    driver_birth = Column('driver_birth', String)
    driver_salary = Column('driver_salary', Integer)
    company_id = Column(Integer, ForeignKey('companies.company_id'))

    __table_args__ = (UniqueConstraint('driver_id', 'driver_name', 'driver_phone',
                                       'driver_email', 'driver_passport', 'driver_license'),)

    def __init__(self, dataframe_table, row_index):
        self.driver_name = dataframe_table['driver_name'][row_index]
        self.driver_phone = dataframe_table['driver_phone'][row_index]
        self.driver_email = dataframe_table['driver_email'][row_index]
        self.driver_passport = dataframe_table['driver_passport'][row_index]
        self.driver_address = dataframe_table['driver_address'][row_index]
        self.driver_license = dataframe_table['driver_license'][row_index]
        self.driver_birth = dataframe_table['driver_birth'][row_index]
        self.driver_salary = dataframe_table['driver_salary'][row_index]
        self.company_id = dataframe_table['company_id'][row_index]

    def __repr__(self):
        return (f"DRIVERS\n"
                f"ID: {self.driver_id},\n"
                f"Name: {self.driver_name},\n"
                f"Phone: {self.driver_phone},\n"
                f"Email: {self.driver_email},\n"
                f"Passport: {self.driver_passport},\n"
                f"Address: {self.driver_address},\n"
                f"License: {self.driver_license},\n"
                f"Date of birth: {self.driver_birth},\n"
                f"Salary: {self.driver_salary},\n"
                f"Company's ID: {self.company_id}\n")

class Trips(Base):

    """
    This class allows to create and manipulate table named 'trips' in the SQL database
    """

    __tablename__ = 'trips'

    trip_id = Column('trip_id', Integer, primary_key=True, nullable=False)
    trip_destination = Column('trip_destination', String, nullable=False)
    trip_date = Column('trip_date', String, nullable=False)
    trip_mileage_m = Column('trip_mileage_m', Integer, nullable=False)
    trip_total_expense = Column('trip_total_expense', Integer, nullable=False)
    trip_allowance = Column('trip_allowance', Integer, nullable=False)
    car_id = Column(Integer, ForeignKey('cars.car_id'))
    driver_id = Column(Integer, ForeignKey('drivers.driver_id'))
    company_id = Column(Integer, ForeignKey('companies.company_id'))

    def __init__(self, dataframe_table, row_index):
        self.trip_destination = dataframe_table['trip_destination'][row_index]
        self.trip_date = dataframe_table['trip_date'][row_index]
        self.trip_mileage_m = dataframe_table['trip_mileage_m'][row_index]
        self.trip_total_expense = dataframe_table['trip_total_expense'][row_index]
        self.trip_allowance = dataframe_table['trip_allowance'][row_index]

    def __repr__(self):
        return (f"TRIPS\n"
                f"ID: {self.trip_id},\n"
                f"Destination: {self.trip_destination},\n"
                f"Date: {self.trip_date},\n"
                f"Mileage: {self.trip_mileage_m},\n"
                f"Total expense: {self.trip_total_expense},\n"
                f"Allowance: {self.trip_allowance},\n"
                f"Car's ID: {self.car_id},\n"
                f"Driver's ID: {self.driver_id},\n"
                f"Company's ID: {self.company_id}\n")

class ShipmentsContent(Base):

    """
    This class allows to create and manipulate table named 'shipments_content' in the SQL database
    """

    __tablename__ = 'shipments_content'

    content_invoice_id = Column('content_invoice_id', Integer, primary_key=True, nullable=False)
    content_invoice_date = Column('content_invoice_date', String, nullable=False)
    content_amount = Column('content_amount', Integer, nullable=False)
    content_price = Column('content_price', DECIMAL, nullable=False)
    content_total = Column('content_total', DECIMAL, nullable=False)
    content_currency = Column('content_currency', String, nullable=False)
    currency_usd_id = Column(Integer, ForeignKey('currency_usd.currency_usd_id'))
    currency_eur_id = Column(Integer, ForeignKey('currency_eur.currency_eur_id'))
    flower_id = Column(Integer, ForeignKey('flower_type.flower_id'))
    supplier_id = Column(Integer, ForeignKey('suppliers.supplier_id'))
    shipment_id = Column(Integer, ForeignKey('shipments.shipment_id'))

    def __init__(self, dataframe_table, row_index):
        self.content_invoice_date = dataframe_table['content_invoice_date'][row_index]
        self.content_amount = dataframe_table['content_amount'][row_index]
        self.content_price = dataframe_table['content_price'][row_index]
        self.content_total = dataframe_table['content_total'][row_index]
        self.content_currency = dataframe_table['content_currency'][row_index]

    def __repr__(self):
        return (f"TRIPS\n"
                f"ID: {self.content_invoice_id},\n"
                f"Invoice date: {self.content_invoice_date},\n"
                f"Amount: {self.content_amount},\n"
                f"Price: {self.content_price},\n"
                f"Total: {self.content_total},\n"
                f"Currency type: {self.content_currency},\n"
                f"USD ID: {self.currency_usd_id},\n"
                f"EUR ID: {self.currency_eur_id},\n"
                f"Flower's ID: {self.flower_id},\n"
                f"Supplier's ID: {self.supplier_id},\n"
                f"Shipment's ID: {self.shipment_id}\n")

class Shipments(Base):

    """
    This class allows to create and manipulate table named 'shipments' in the SQL database
    """

    __tablename__ = 'shipments'

    shipment_id = Column('shipment_id', Integer, primary_key=True, nullable=False)
    shipment_date = Column('shipment_date', Date, nullable=False)
    shipment_box_amount = Column('shipment_box_amount', Integer, nullable=False)
    shipment_box_full = Column('shipment_box_full', DECIMAL, nullable=False)
    shipment_weight_fact = Column('shipment_weight_fact', Integer)
    shipment_weight_vol = Column('shipment_weight_vol', Integer)
    shipment_volume = Column('shipment_volume', DECIMAL, nullable=False)
    shipment_marking = Column('shipment_marking', String, nullable=False)
    shipment_awb = Column('shipment_awb', String)
    shipment_country = Column('shipment_country', String, nullable=False)
    shipment_supplier = Column('shipment_supplier', String, nullable=False)
    shipment_ams_arrival = Column('shipment_ams_arrival', String)
    shipment_msc_arrival = Column('shipment_msc_arrival', String)
    shipment_rstv_arrival = Column('shipment_rstv_arrival', String)
    shipment_krd_arrival = Column('shipment_krd_arrival', String)
    shipment_arm_arrival = Column('shipment_arm_arrival', String)
    shipment_status = Column('shipment_status', String)
    shipment_truck_name = Column('shipment_truck_name', String, nullable=False)
    shipment_truck_balance = Column('shipment_truck_balance', String)
    shipment_forever_balance = Column('shipment_forever_balance', String, nullable=False)
    shipment_comment = Column('shipment_comment', String)
    customer_id = Column(Integer, ForeignKey('customers.customer_id'))
    expense_forever_id = Column(Integer, ForeignKey('operations_forever.expense_forever_id'))
    expense_ip_id = Column(Integer, ForeignKey('operations_iphandlers.expense_ip_id'))
    sale_id = Column(Integer, ForeignKey('operations_sales.sale_id'))
    car_id = Column(Integer, ForeignKey('cars.car_id'))
    driver_id = Column(Integer, ForeignKey('drivers.driver_id'))
    supplier_id = Column(Integer, ForeignKey('suppliers.supplier_id'))
    content_invoice_id = Column(Integer, ForeignKey('shipments_content.content_invoice_id'))
    company_id = Column(Integer, ForeignKey('companies.company_id'))

    def __init__(self, dataframe_table, row_index):
        self.shipment_box_amount = dataframe_table['shipment_box_amount'][row_index]
        self.shipment_date = datetime.strptime(dataframe_table['shipment_date'][row_index], '%Y-%m-%d').date()
        self.shipment_box_full = dataframe_table['shipment_box_full'][row_index]
        self.shipment_weight_fact = dataframe_table['shipment_weight_fact'][row_index]
        self.shipment_weight_vol = dataframe_table['shipment_weight_vol'][row_index]
        self.shipment_volume = dataframe_table['shipment_volume'][row_index]
        self.shipment_marking = dataframe_table['shipment_marking'][row_index]
        self.shipment_awb = dataframe_table['shipment_awb'][row_index]
        self.shipment_country = dataframe_table['shipment_country'][row_index]
        self.shipment_supplier = dataframe_table['shipment_supplier'][row_index]
        self.shipment_truck_name = dataframe_table['shipment_truck_name'][row_index]
        self.shipment_forever_balance = dataframe_table['shipment_forever_balance'][row_index]
        self.shipment_comment = dataframe_table['shipment_comment'][row_index]
        self.shipment_status = dataframe_table['shipment_status'][row_index]
        self.shipment_truck_balance = dataframe_table['shipment_truck_balance'][row_index]

    def __repr__(self):
        return (f"SHIPMENTS\n"
                f"ID: {self.shipment_id},\n"
                f"Date: {self.shipment_date},\n"
                f"Boxes: {self.shipment_box_amount},\n"
                f"Full boxes: {self.shipment_box_full},\n"
                f"Weight factual: {self.shipment_weight_fact},\n"
                f"Weight gross: {self.shipment_weight_vol},\n"
                f"Volume: {self.shipment_volume},\n"
                f"Marking: {self.shipment_marking},\n"
                f"Awb: {self.shipment_awb},\n"
                f"Country: {self.shipment_country},\n"
                f"Supplier: {self.shipment_supplier},\n"
                f"Amsterdam arrival time: {self.shipment_ams_arrival},\n"
                f"Moscow arrival time: {self.shipment_msc_arrival},\n"
                f"Rostov-on-Don arrival time: {self.shipment_rstv_arrival},\n"
                f"Krasnodar arrival time: {self.shipment_krd_arrival},\n"
                f"Armavir arrival time: {self.shipment_arm_arrival},\n"
                f"Status: {self.shipment_status},\n"
                f"Truck name: {self.shipment_truck_name},\n"
                f"Truck balance: {self.shipment_truck_balance},\n"
                f"Forever sub balance: {self.shipment_forever_balance},\n"
                f"Comment: {self.shipment_comment},\n"
                f"Expense Forever ID: {self.expense_forever_id},\n"
                f"Expense Iphandlers ID: {self.expense_ip_id},\n"
                f"Sale ID: {self.sale_id},\n"
                f"Customer's ID: {self.customer_id},\n"
                f"Car ID: {self.car_id},\n"
                f"Driver's ID: {self.driver_id},\n"
                f"Supplier's ID: {self.supplier_id},\n"
                f"Invoice ID: {self.content_invoice_id},\n"
                f"Companie's ID: {self.company_id}\n")

    def update_manual_values(self, dataframe_table, row_index):

        """
        This function fills manually updated values in the table 'shipments'
        """

        self.shipment_ams_arrival = dataframe_table['shipment_ams_arrival'][row_index]
        self.shipment_msc_arrival = dataframe_table['shipment_msc_arrival'][row_index]
        self.shipment_rstv_arrival = dataframe_table['shipment_rstv_arrival'][row_index]
        self.shipment_krd_arrival = dataframe_table['shipment_krd_arrival'][row_index]
        self.shipment_arm_arrival = dataframe_table['shipment_arm_arrival'][row_index]
        self.shipment_status = dataframe_table['shipment_status'][row_index]
        self.shipment_truck_balance = dataframe_table['shipment_truck_balance'][row_index]

class OperationsSales(Base):

    """
    This class allows to create and manipulate table named 'operations_sales' in the SQL database
    """

    __tablename__ = 'operations_sales'

    sale_id = Column('sale_id', Integer, primary_key=True, nullable=False)
    sale_date = Column('sale_date', Date, nullable=False)
    sale_type = Column('sale_type', String, nullable=False)
    sale_marking = Column('sale_marking', String, nullable=False)
    sale_full_box = Column('sale_full_box', DECIMAL)
    sale_customer = Column('sale_customer', String)
    sale_content_supplier = Column('sale_content_supplier', String, nullable=False)
    sale_total_usd = Column('sale_total_usd', DECIMAL, nullable=False)
    sale_total_eur = Column('sale_total_eur', DECIMAL, nullable=False)
    sale_total_rub = Column('sale_total_rub', DECIMAL)
    sale_currency = Column('sale_currency', String, nullable=False)
    sale_currency_rate = Column('sale_currency_rate', DECIMAL)
    sale_volume = Column('sale_volume', DECIMAL, nullable=False)
    sale_country = Column('sale_country', String)
    sale_weight = Column('sale_weight', Integer, nullable=False)
    sale_awb = Column('sale_awb', Integer, nullable=False)
    sale_currency_percent = Column('sale_currency_percent', DECIMAL)
    sale_currency_markup = Column('sale_currency_markup', DECIMAL)
    sale_total_percent = Column('sale_total_percent', DECIMAL)
    sale_total_markup = Column('sale_total_markup', Integer)
    sale_price_kg = Column('sale_price_kg', DECIMAL)
    sale_price_pallet = Column('sale_price_pallet', Integer)
    sale_price_troll = Column('sale_price_troll', Integer)
    expense_forever_id = Column(Integer, ForeignKey('operations_forever.expense_forever_id'))
    expense_ip_id = Column(Integer, ForeignKey('operations_iphandlers.expense_ip_id'))
    currency_usd_id = Column(Integer, ForeignKey('currency_usd.currency_usd_id'))
    currency_eur_id = Column(Integer, ForeignKey('currency_eur.currency_eur_id'))
    customer_id = Column(Integer, ForeignKey('customers.customer_id'))
    shipment_id = Column(Integer, ForeignKey('shipments.shipment_id'))
    content_invoice_id = Column(Integer, ForeignKey('shipments_content.content_invoice_id'))
    manager_id = Column(Integer, ForeignKey('managers.manager_id'))
    company_id = Column(Integer, ForeignKey('companies.company_id'))
    supplier_id  = Column(Integer, ForeignKey('suppliers.supplier_id'))

    def __init__(self, dataframe_table, row_index):
        self.sale_date = datetime.strptime(dataframe_table['sale_date'][row_index], '%Y-%m-%d').date()
        self.sale_type = dataframe_table['sale_type'][row_index]
        self.sale_marking = dataframe_table['sale_marking'][row_index]
        self.sale_content_supplier = dataframe_table['sale_content_supplier'][row_index]
        self.sale_total_usd = dataframe_table['sale_total_usd'][row_index]
        self.sale_total_eur = dataframe_table['sale_total_eur'][row_index]
        self.sale_currency = dataframe_table['sale_currency'][row_index]
        self.sale_volume = dataframe_table['sale_volume'][row_index]
        self.sale_weight = dataframe_table['sale_weight'][row_index]
        self.sale_awb = dataframe_table['sale_awb'][row_index]
        self.supplier_id = dataframe_table['supplier_id'][row_index]
        self.sale_full_box = dataframe_table['sale_full_box'][row_index]

    def __repr__(self):
        return (f"SALES\n"
                f"ID: {self.sale_id},\n"
                f"Date: {self.sale_date},\n"
                f"Type: {self.sale_type},\n"
                f"Marking: {self.sale_marking},\n"
                f"Customer: {self.sale_customer},\n"
                f"Supplier: {self.sale_content_supplier},\n"
                f"Total/USD: {self.sale_total_usd},\n"
                f"Total/EUR: {self.sale_total_eur},\n"
                f"Total/RUB: {self.sale_total_rub},\n"
                f"Currency type: {self.sale_currency},\n"
                f"Currency rate: {self.sale_currency_rate},\n"
                f"Full box: {self.sale_full_box},\n"                
                f"Volume: {self.sale_volume},\n"
                f"Country: {self.sale_country},\n"
                f"Weight: {self.sale_weight},\n"
                f"Awb: {self.sale_awb},\n"
                f"Currency + %: {self.sale_currency_percent},\n"
                f"Currency + RUB: {self.sale_currency_markup},\n"
                f"Total + %: {self.sale_total_percent},\n"
                f"Total + RUB: {self.sale_total_markup},\n"
                f"Price/kg: {self.sale_price_kg},\n"
                f"Price/palett: {self.sale_price_pallet},\n"
                f"Price/troll: {self.sale_price_troll},\n"
                f"Expense Forever ID: {self.expense_forever_id},\n"
                f"Expense Iphandlers ID: {self.expense_ip_id},\n"
                f"USD ID: {self.currency_usd_id},\n"
                f"EUR ID: {self.currency_eur_id},\n"
                f"Customer's ID: {self.customer_id},\n"
                f"Shipment's ID: {self.shipment_id},\n"
                f"Invoice ID: {self.content_invoice_id},\n"
                f"Manager's ID: {self.manager_id},\n"
                f"Company's ID: {self.company_id},\n"
                f"Supplier's ID: {self.supplier_id}\n")

class OperationsForever(Base):

    """
    This class allows to create and manipulate table named 'operations_forever' in the SQL database
    """

    __tablename__ = 'operations_forever'

    expense_forever_id = Column('expense_forever_id', Integer, primary_key=True, nullable=False)
    expense_customer = Column('expense_customer', String)
    expense_date = Column('expense_date', Date, nullable=False)
    expense_type = Column('expense_type', String, nullable=False)
    expense_total_usd = Column('expense_total_usd', DECIMAL, nullable=False)
    expense_total_eur = Column('expense_total_eur', DECIMAL, nullable=False)
    expense_total_rub = Column('expense_total_rub', DECIMAL, nullable=False)
    expense_country = Column('expense_country', String)
    expense_currency = Column('expense_currency', String, nullable=False)
    expense_currency_rate = Column('expense_currency_rate', DECIMAL, nullable=False)
    expense_awb = Column('expense_awb', String, nullable =False)
    expense_content_supplier = Column('expense_content_supplier', String, nullable=False)
    expense_marking = Column('expense_marking', String, nullable=False)
    expense_full_box = Column('expense_full_box', DECIMAL, nullable=False)
    expense_weight = Column('expense_weight', Integer, nullable=False)
    expense_volume = Column('expense_volume', DECIMAL, nullable=False)
    expense_balance_code = Column('expense_balance_code', Integer)
    expense_balance_currency = Column('expense_balance_currency', String)
    content_invoice_id = Column(Integer, ForeignKey('shipments_content.content_invoice_id'))
    supplier_id  = Column(Integer, ForeignKey('suppliers.supplier_id'))
    shipment_id = Column(Integer, ForeignKey('shipments.shipment_id'))
    car_id = Column(Integer, ForeignKey('cars.car_id'))
    customer_id = Column(Integer, ForeignKey('customers.customer_id'))
    company_id = Column(Integer, ForeignKey('companies.company_id'))
    trip_id = Column(Integer, ForeignKey('trips.trip_id'))
    driver_id = Column(Integer, ForeignKey('drivers.driver_id'))
    manager_id = Column(Integer, ForeignKey('managers.manager_id'))
    sale_id = Column(Integer, ForeignKey('operations_sales.sale_id'))

    def __init__(self, dataframe_table, row_index):
        self.expense_date = datetime.strptime(dataframe_table['expense_date'][row_index], '%Y-%m-%d').date()
        self.expense_type = dataframe_table['expense_type'][row_index]
        self.expense_total_usd = dataframe_table['expense_total_usd'][row_index]
        self.expense_total_eur = dataframe_table['expense_total_eur'][row_index]
        self.expense_total_rub = dataframe_table['expense_total_rub'][row_index]
        self.expense_currency = dataframe_table['expense_currency'][row_index]
        self.expense_currency_rate = dataframe_table['expense_currency_rate'][row_index]
        self.expense_awb = dataframe_table['expense_awb'][row_index]
        self.expense_content_supplier = dataframe_table['expense_content_supplier'][row_index]
        self.expense_marking = dataframe_table['expense_marking'][row_index]
        self.expense_full_box = dataframe_table['expense_full_box'][row_index]
        self.expense_weight = dataframe_table['expense_weight'][row_index]
        self.expense_volume = dataframe_table['expense_volume'][row_index]
        self.expense_balance_code = dataframe_table['expense_balance_code'][row_index]
        self.expense_balance_currency = dataframe_table['expense_balance_currency'][row_index]
        self.supplier_id  = dataframe_table['supplier_id'][row_index]

    def __repr__(self):
        return (f"EXPENSES FOREVER\n"
                f"ID: {self.expense_forever_id},\n"
                f"Customer: {self.expense_customer},\n"
                f"Date: {self.expense_date},\n"
                f"Type: {self.expense_type},\n"
                f"Total/USD: {self.expense_total_usd},\n"
                f"Total/EUR: {self.expense_total_eur},\n"
                f"Total/RUB: {self.expense_total_rub},\n"
                f"Country: {self.expense_country},\n"
                f"Currency type: {self.expense_currency},\n"
                f"Currency rate: {self.expense_currency_rate},\n"
                f"Awb: {self.expense_awb},\n"
                f"Content's supplier: {self.expense_content_supplier},\n"
                f"Marking: {self.expense_marking},\n"
                f"Full boxes: {self.expense_full_box},\n"
                f"Weight: {self.expense_weight},\n"
                f"Volume: {self.expense_volume},\n"
                f"Balance code: {self.expense_balance_code},\n"
                f"Balance currency type: {self.expense_balance_currency},\n"
                f"Invoice ID: {self.content_invoice_id},\n"
                f"Supplier's ID: {self.supplier_id},\n"
                f"Shipment's ID: {self.shipment_id},\n"
                f"Car ID: {self.car_id},\n"
                f"Customer's ID: {self.customer_id},\n"
                f"Company's ID: {self.company_id},\n"
                f"Trip ID: {self.trip_id},\n"
                f"Driver's ID: {self.driver_id},\n"
                f"Manager's ID: {self.manager_id},\n"
                f"Sale ID: {self.sale_id}\n")

class OperationsIphandlers(Base):

    """
    This class allows to create and manipulate table named
    'operations_iphandlers' in the SQL database
    """

    __tablename__ = 'operations_iphandlers'

    expense_ip_id = Column('expense_ip_id', Integer, primary_key=True, nullable=False)
    expense_customer = Column('expense_customer', String)
    expense_eta_date = Column('expense_eta_date', Date, nullable=False)
    expense_load_date = Column('expense_load_date', String, nullable=False)
    expense_total = Column('expense_total', DECIMAL, nullable=False)
    expense_country = Column('expense_country', String)
    expense_currency = Column('expense_currency', String, nullable=False)
    expense_awb = Column('expense_awb', String, nullable =False)
    expense_marking = Column('expense_marking', String, nullable=False)
    expense_box = Column('expense_box', Integer, nullable=False)
    expense_full_box = Column('expense_full_box', DECIMAL, nullable=False)
    expense_weight = Column('expense_weight', Integer, nullable=False)
    expense_account = Column('expense_account', String, nullable=False)
    supplier_id  = Column(Integer, ForeignKey('suppliers.supplier_id'))
    shipment_id = Column(Integer, ForeignKey('shipments.shipment_id'), unique=True)
    expense_forever_id = Column(Integer, ForeignKey('operations_forever.expense_forever_id'))
    customer_id = Column(Integer, ForeignKey('customers.customer_id'))
    company_id = Column(Integer, ForeignKey('companies.company_id'))
    sale_id = Column(Integer, ForeignKey('operations_sales.sale_id'))

    def __init__(self, dataframe_table, row_index):
        self.expense_eta_date = datetime.strptime(dataframe_table['expense_eta_date'][row_index], '%Y-%m-%d').date()
        self.expense_load_date = dataframe_table['expense_load_date'][row_index]
        self.expense_account = dataframe_table['expense_account'][row_index]
        self.expense_total = dataframe_table['expense_total'][row_index]
        self.expense_currency = 'eur'
        self.expense_awb = dataframe_table['expense_awb'][row_index]
        self.expense_marking = dataframe_table['expense_marking'][row_index]
        self.expense_box = dataframe_table['expense_box'][row_index]
        self.expense_full_box = dataframe_table['expense_full_box'][row_index]
        self.expense_weight = dataframe_table['expense_weight'][row_index]

    def __repr__(self):
        return (f"EXPENSES IPHANDLERS\n"
                f"ID: {self.expense_ip_id},\n"
                f"Customer: {self.expense_customer},\n"
                f"ETA Date: {self.expense_eta_date},\n"
                f"Loading Date: {self.expense_load_date},\n"
                f"Accountable: {self.expense_account},\n"                
                f"Total: {self.expense_total},\n"
                f"Country: {self.expense_country},\n"
                f"Currency type: {self.expense_currency},\n"
                f"Awb: {self.expense_awb},\n"
                f"Marking: {self.expense_marking},\n"
                f"Boxes: {self.expense_box},\n"                
                f"Full boxes: {self.expense_full_box},\n"
                f"Weight: {self.expense_weight},\n"
                f"Expense Forever ID: {self.expense_forever_id},\n"
                f"Supplier's ID: {self.supplier_id},\n"
                f"Shipment's ID: {self.shipment_id},\n"
                f"Customer's ID: {self.customer_id},\n"
                f"Company's ID: {self.company_id},\n"
                f"Sale ID: {self.sale_id}\n")

def create_table(dataframe_intro, table_class, session_current):

    """
    This function creates a table based on the Class and dataframe provided

    """

    queries = []

    # Check if it is a duplicate of a record
    for ind in dataframe_intro.index:

        result_duplicate = ''

        if table_class == OperationsForever:

            query_duplicate = session_current.query(OperationsForever).filter(and_(
                              OperationsForever.expense_awb == dataframe_intro
                                                               .at[ind, 'expense_awb'],
                              OperationsForever.expense_marking == dataframe_intro
                                                                   .at[ind, 'expense_marking'],
                              OperationsForever.expense_full_box == dataframe_intro
                                                                    .at[ind, 'expense_full_box'],
                              OperationsForever.expense_date == dataframe_intro
                                                                .at[ind, 'expense_date']
                                                                ))

            result_duplicate = query_duplicate.first()

        elif table_class == OperationsIphandlers:

            query_duplicate = session_current.query(OperationsIphandlers).filter(and_(
                              OperationsIphandlers.expense_awb == dataframe_intro
                                                                  .at[ind, 'expense_awb'],
                              OperationsIphandlers.expense_marking == dataframe_intro
                                                                      .at[ind, 'expense_marking'],
                              OperationsIphandlers.expense_full_box == dataframe_intro
                                                                       .at[ind, 'expense_full_box'],
                              OperationsIphandlers.expense_eta_date == dataframe_intro
                                                                       .at[ind, 'expense_eta_date']
                                                                       ))

            result_duplicate = query_duplicate.first()

        elif table_class == OperationsSales:

            query_duplicate = session_current.query(OperationsSales).filter(and_(
                              OperationsSales.sale_awb == dataframe_intro
                                                          .at[ind, 'sale_awb'],
                              OperationsSales.sale_marking == dataframe_intro
                                                              .at[ind, 'sale_marking'],
                              OperationsSales.sale_full_box == dataframe_intro
                                                               .at[ind, 'sale_full_box'],
                              OperationsSales.sale_date == dataframe_intro
                                                           .at[ind, 'sale_date']
                                                           ))

            result_duplicate = query_duplicate.first()

        elif table_class == Shipments:

            query_duplicate = session_current.query(Shipments).filter(and_(
                              Shipments.shipment_awb == dataframe_intro
                                                        .at[ind, 'shipment_awb'],
                              Shipments.shipment_marking == dataframe_intro
                                                            .at[ind, 'shipment_marking'],
                              Shipments.shipment_box_full == dataframe_intro
                                                             .at[ind, 'shipment_box_full'],
                              Shipments.shipment_truck_name == dataframe_intro
                                                               .at[ind, 'shipment_truck_name'],
                              Shipments.shipment_weight_vol == dataframe_intro
                                                               .at[ind, 'shipment_weight_vol']
                                                               ))

            result_duplicate = query_duplicate.first()

        elif table_class == Markings:

            query_duplicate = session_current.query(Markings).filter(
                              Markings.marking_name == dataframe_intro
                                                        .at[ind, 'marking_name']
                                                        )

            result_duplicate = query_duplicate.first()

        if (result_duplicate is None or
            result_duplicate == '' or
            result_duplicate == 'None'):

            table = table_class(dataframe_intro, ind)

            session_current.add(table)

            session_current.commit()

    # Fill the queries based on which table is being created
    if table_class == OperationsForever:

        # Dictionary of queries
        queries = [update(OperationsForever).values(customer_id=session_current
                    .query(Markings.customer_id)
                    .filter(Markings.marking_name == OperationsForever.expense_marking)
                    .as_scalar())
                    .where(OperationsForever.customer_id.is_(None)),

                    update(OperationsForever).values(expense_customer=session_current
                    .query(Customers.customer_name)
                    .filter(Customers.customer_id == OperationsForever.customer_id)
                    .as_scalar())
                    .where(OperationsForever.expense_customer.is_(None)),

                    update(OperationsForever).values(shipment_id=session_current
                    .query(Shipments.shipment_id)
                    .filter(and_(
                        Shipments.expense_forever_id.is_(None),
                        Shipments.shipment_awb == OperationsForever.expense_awb,
                        Shipments.shipment_box_full == OperationsForever.expense_full_box,
                        Shipments.shipment_marking == OperationsForever.expense_marking
                    ))
                    .as_scalar())
                    .where(OperationsForever.shipment_id.is_(None)),

                    update(OperationsForever).values(shipment_id=session_current
                    .query(Shipments.shipment_id)
                    .filter(and_(
                        Shipments.expense_forever_id.is_(None),
                        Shipments.shipment_marking == OperationsForever.expense_marking,
                        Shipments.shipment_awb == OperationsForever.expense_awb))
                    .as_scalar())
                    .where(OperationsForever.shipment_id.is_(None)),

                    update(Shipments).values(expense_forever_id=session_current
                    .query(OperationsForever.expense_forever_id)
                    .filter(OperationsForever.shipment_id == Shipments.shipment_id)
                    .as_scalar())
                    .where(Shipments.expense_forever_id.is_(None)),

                    update(OperationsForever).values(expense_country=session_current
                    .query(Shipments.shipment_country)
                    .filter(
                        Shipments.shipment_id == OperationsForever.shipment_id)
                    .as_scalar())
                    .where(OperationsForever.expense_country.is_(None))
                    ]

    elif table_class == OperationsIphandlers:

        # Dictionary of queries
        queries = [update(OperationsIphandlers).values(customer_id=session_current
                    .query(Markings.customer_id)
                    .filter(Markings.marking_name == OperationsIphandlers.customer_id)
                    .as_scalar())
                    .where(OperationsIphandlers.customer_id.is_(None)),

                    update(OperationsIphandlers).values(expense_customer=session_current
                    .query(Customers.customer_name)
                    .filter(Customers.customer_id == OperationsIphandlers.customer_id)
                    .as_scalar())
                    .where(OperationsIphandlers.expense_customer.is_(None)),

                    update(OperationsIphandlers).values(shipment_id=session_current
                    .query(Shipments.shipment_id)
                    .filter(and_(
                        Shipments.expense_ip_id.is_(None),
                        Shipments.shipment_awb == OperationsIphandlers.expense_awb,
                        Shipments.shipment_marking == OperationsIphandlers.expense_marking
                    ))
                    .as_scalar())
                    .where(OperationsIphandlers.shipment_id.is_(None)),

                    update(Shipments).values(expense_ip_id=session_current
                    .query(OperationsIphandlers.expense_ip_id)
                    .filter(OperationsIphandlers.shipment_id == Shipments.shipment_id)
                    .as_scalar())
                    .where(Shipments.expense_ip_id.is_(None))
                    ]

    elif table_class == OperationsSales:

        # Dictionary of OPERATIONS_SALES queries
        queries = [update(OperationsSales).values(customer_id=session_current
                    .query(Markings.customer_id)
                    .filter(Markings.marking_name == OperationsSales.sale_marking)
                    .as_scalar())
                    .where(OperationsSales.customer_id.is_(None)),

                    update(OperationsSales).values(sale_customer=session_current
                    .query(Customers.customer_name)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(OperationsSales.sale_customer.is_(None)),

                    update(OperationsSales).values(shipment_id=session_current
                    .query(OperationsForever.shipment_id)
                    .filter(and_(
                        OperationsForever.expense_awb == OperationsSales.sale_awb,
                        OperationsForever.expense_full_box == OperationsSales.sale_full_box,
                        OperationsForever.expense_marking == OperationsSales.sale_marking
                    ))
                    .as_scalar())
                    .where(OperationsSales.shipment_id.is_(None)),

                    update(OperationsSales).values(expense_forever_id=session_current
                    .query(OperationsForever.expense_forever_id)
                    .filter(and_(
                        OperationsForever.expense_awb == OperationsSales.sale_awb,
                        OperationsForever.expense_full_box == OperationsSales.sale_full_box,
                        OperationsForever.expense_marking == OperationsSales.sale_marking,
                        OperationsSales.supplier_id == 2
                    ))
                    .as_scalar())
                    .where(OperationsSales.expense_forever_id.is_(None)),

                    update(OperationsSales).values(expense_ip_id=session_current
                    .query(OperationsIphandlers.expense_ip_id)
                    .filter(and_(
                        OperationsIphandlers.expense_awb == OperationsSales.sale_awb,
                        OperationsIphandlers.expense_full_box == OperationsSales.sale_full_box,
                        OperationsIphandlers.expense_marking == OperationsSales.sale_marking,
                        OperationsSales.supplier_id == 5
                    ))
                    .as_scalar())
                    .where(OperationsSales.expense_ip_id.is_(None)),

                    update(OperationsSales).values(sale_country=session_current
                    .query(Shipments.shipment_country)
                    .filter(
                        Shipments.shipment_id == OperationsSales.shipment_id)
                    .as_scalar())
                    .where(OperationsSales.sale_country.is_(None)),

                    update(OperationsSales).values(sale_currency_markup=session_current
                    .query(Customers.customer_dollar_trans_rate)
                    .filter(
                        Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_currency_markup.is_(None),
                        OperationsSales.sale_currency == 'usd')),

                    update(OperationsSales).values(sale_currency_markup=session_current
                    .query(Customers.customer_euro_trans_rate)
                    .filter(
                        Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_currency_markup.is_(None),
                        OperationsSales.sale_currency == 'eur')),

                    update(OperationsSales).values(sale_price_kg=session_current
                    .query(Customers.customer_cons_kg)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_kg.is_(None),
                        OperationsSales.sale_type.like("%консолидат%"),
                        OperationsSales.sale_country == 'ec')),

                    update(OperationsSales).values(sale_price_pallet=session_current
                    .query(Customers.customer_holl_pallet)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_pallet.is_(None),
                        OperationsSales.sale_type.like("%срезка%"),
                        OperationsSales.sale_country == 'den')),

                    update(OperationsSales).values(sale_price_troll=session_current
                    .query(Customers.customer_troll)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_troll.is_(None),
                        OperationsSales.sale_type.like("%телег%"),
                        OperationsSales.sale_country == 'den')),

                    update(OperationsSales).values(sale_price_kg=session_current
                    .query(Customers.customer_eq_kg)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_kg.is_(None),
                        OperationsSales.sale_type.like("%срезка%"),
                        OperationsSales.sale_country == 'ec')),

                    update(OperationsSales).values(sale_price_kg=session_current
                    .query(Customers.customer_col_kg)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_kg.is_(None),
                        OperationsSales.sale_type.like("%срезка%"),
                        OperationsSales.sale_country == 'co')),

                    update(OperationsSales).values(sale_price_kg=session_current
                    .query(Customers.customer_ken_kg)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_kg.is_(None),
                        OperationsSales.sale_type.like("%срезка%"),
                        OperationsSales.sale_country == 'ke')),

                    update(OperationsSales).values(sale_price_pallet=session_current
                    .query(Customers.customer_holl_pallet)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_pallet.is_(None),
                        OperationsSales.sale_type.like("%срезка%"),
                        OperationsSales.sale_country == 'nl')),

                    update(OperationsSales).values(sale_price_troll=session_current
                    .query(Customers.customer_troll)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_troll.is_(None),
                        OperationsSales.sale_type.like("%телег%"),
                        OperationsSales.sale_country == 'nl')),

                    update(OperationsSales).values(sale_price_kg=session_current
                    .query(Customers.customer_isr_kg)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_kg.is_(None),
                        OperationsSales.sale_type.like("%весу%"),
                        OperationsSales.sale_country == 'il')),

                    update(OperationsSales)
                    .values(sale_price_pallet=session_current
                    .query(Customers.customer_isr_pallet)
                    .filter(Customers.customer_id == OperationsSales.customer_id)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_price_pallet.is_(None),
                        OperationsSales.sale_type.like("%объему%"),
                        OperationsSales.sale_country == 'il')),

                    update(OperationsSales)
                    .values(sale_currency_rate=session_current
                    .query(CurrencyUsd.currency_rate)
                    .filter(CurrencyUsd.currency_date == OperationsSales.sale_date)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_currency_rate.is_(None),
                        OperationsSales.sale_currency == 'usd')),

                    update(OperationsSales)
                    .values(sale_currency_rate=session_current
                    .query(CurrencyEur.currency_rate)
                    .filter(CurrencyEur.currency_date == OperationsSales.sale_date)
                    .as_scalar())
                    .where(and_(
                        OperationsSales.sale_currency_rate.is_(None),
                        OperationsSales.sale_currency == 'eur')),

                    update(OperationsSales)
                    .values(sale_total_rub=OperationsSales.sale_weight * \
                            OperationsSales.sale_price_kg * \
                            (OperationsSales.sale_currency_rate + \
                             OperationsSales.sale_currency_markup))
                    .where(and_(
                        OperationsSales.sale_total_rub.is_(None),
                        OperationsSales.sale_price_kg.is_not(None),
                        OperationsSales.sale_currency == 'usd')),

                    update(OperationsSales)
                    .values(sale_total_rub=OperationsSales.sale_volume * \
                            OperationsSales.sale_price_pallet * \
                            (OperationsSales.sale_currency_rate + \
                             OperationsSales.sale_currency_markup))
                    .where(and_(
                        OperationsSales.sale_total_rub.is_(None),
                        OperationsSales.sale_price_pallet.is_not(None),
                        OperationsSales.sale_currency == 'usd')),

                    update(OperationsSales)
                    .values(sale_total_rub=OperationsSales.sale_full_box * \
                            OperationsSales.sale_price_troll * \
                            (OperationsSales.sale_currency_rate + \
                             OperationsSales.sale_currency_markup))
                    .where(and_(
                        OperationsSales.sale_total_rub.is_(None),
                        OperationsSales.sale_price_troll.is_not(None),
                        OperationsSales.sale_currency == 'usd'))
                    ]

    elif table_class == Shipments:

        # Dictionary of SHIPMENTS queries
        queries = [update(Shipments).values(customer_id=session_current
                    .query(Markings.customer_id)
                    .filter(Markings.marking_name == Shipments.shipment_marking)
                    .as_scalar())
                    ]

    elif table_class == Markings:

        queries = [update(Markings).values(customer_id=session_current
                                   .query(Customers.customer_id)
                                   .filter(Customers.customer_name==(
                                           Markings.marking_customer))
                                   .as_scalar())
                                   .where(Markings.customer_id.is_(None))
                    ]

    if queries:

        for query in queries:

            # Execute the update statement
            session_current.execute(query)

            # Commit the changes
            session_current.commit()

    results = session_current.query(table_class).all()

    # Save the original sys.stdout as log_history.txt
    original_stdout = sys.stdout

    file_path = 'C:/autocargo_reports/export/log_history.txt'

    # Check if log file path is valid
    if os.path.exists(file_path):

        with open(file_path, 'a', encoding='utf_8') as file:
            sys.stdout = file

            for r in results:
                print(r)

    # Check if log file path is valid
    else:

        print('Log history file was not found!')
        print('Records were not updated.')

    # Restore the original sys.stdout
    sys.stdout = original_stdout

def database_create(database_path):

    """
    This function creates a table in a database.

    """

    # Create the database and tables
    engine = create_engine(f"sqlite:///{database_path}", echo = False)

    Base.metadata.create_all(bind=engine)

    session = sessionmaker(bind=engine)

    session_current = session()

    # Dictionary of standard tables path + table name
    standard_tables = {Companies:"C:/autocargo_reports/starter/companies.xlsx",
                        Customers:"C:/autocargo_reports/starter/customers_info.xlsx",
                        Suppliers:"C:/autocargo_reports/starter/suppliers.xlsx",
                        BoxType:"C:/autocargo_reports/starter/box_types.xlsx",
                        FlowerType:"C:/autocargo_reports/starter/flower_type.xlsx",
                        Managers:"C:/autocargo_reports/starter/managers.xlsx",
                        Cars:"C:/autocargo_reports/starter/cars.xlsx",
                        Drivers:"C:/autocargo_reports/starter/drivers.xlsx",
                        Markings:"C:/autocargo_reports/starter/markings.xlsx"}

    # Fill standard tables
    for table, path in standard_tables.items():

        if os.path.exists(path):

            dataframe_file = pd.read_excel(path)

            create_table(dataframe_file, table, session_current)

        else:

            print('Standard tables path for table creation is not valid!')

def empty_customer_id(session_current):

    """
    This function fills an empty customer_id of tables

    """

    skip_id = []

    quit_for_loop = False

    cus_id_queries = {Shipments:[Shipments.customer_id,
                                 Shipments.shipment_marking,
                                 Shipments.shipment_id],
                      OperationsForever:[OperationsForever.customer_id,
                                         OperationsForever.expense_marking,
                                         OperationsForever.expense_forever_id],
                      OperationsIphandlers:[OperationsIphandlers.customer_id,
                                            OperationsIphandlers.expense_marking,
                                            OperationsIphandlers.expense_ip_id],
                      OperationsSales:[OperationsSales.customer_id,
                                       OperationsSales.sale_marking,
                                       OperationsSales.sale_id]}

    for table, condition_list in cus_id_queries.items():

        # Condition to identify records where customer_id is None
        condition = condition_list[0].is_(None)

        # Fetch the records that match the empty condition
        records_to_update = session_current.query(table).filter(condition).first()

        while records_to_update:

            if table == Shipments:

                record_names = {table:[records_to_update.shipment_marking,
                                       records_to_update.shipment_id]}

            elif table == OperationsIphandlers:

                record_names = {table:[records_to_update.expense_marking,
                                       records_to_update.expense_ip_id]}

            elif table == OperationsForever:

                if (records_to_update.expense_full_box == 0E-10 and
                    records_to_update.expense_volume == 0E-10):

                    records_to_update.customer_id = '0'

                    session_current.commit()

                    records_to_update = session_current.query(table).filter(condition).first()

                    continue

                record_names = {table:[records_to_update.expense_marking,
                                       records_to_update.expense_forever_id]}

            elif table == OperationsSales:

                record_names = {table:[records_to_update.sale_marking,
                                       records_to_update.sale_id]}

            new_id = ''

            upload = 'no'

            print(border_line)
            print(f"Updating record:\n{records_to_update}.")
            print("Input a customer_id number to link, skip, next or exit")

            match = session_current.query(Markings.customer_id,
                                          Markings.marking_customer,
                                          Markings.marking_name).filter(
                                          Markings.marking_name.in_([record_names[table][0]])
            )

            update_id = update(table).values(customer_id=session_current
                                     .query(Markings.customer_id)
                                     .filter(Markings.marking_name == condition_list[1])
                                     .as_scalar()
                                     ).where(condition_list[1] == record_names[table][0])

            for row in match:

                print(border_line)
                print(f"\nFound a potential match:\nCustomer's ID:{row.customer_id}")
                print(f"Customer's name:{row.marking_customer}")
                print(f"Marking's name:{row.marking_name}\n")

            print(border_line)

            while upload != 'yes':

                new_id = input("\nYour input:\n")

                if (new_id == 'skip' or
                    new_id == 'exit' or
                    new_id == 'next'):

                    if new_id == 'skip':

                        skip_id.append(int(record_names[table][1]))

                    print('Skipping the record...')

                    break

                customer_name = session_current.query(Markings.marking_customer).filter(
                                                    Markings.customer_id.is_(new_id)).first()

                customer_address = session_current.query(Markings.marking_customer_address).filter(
                                                    Markings.customer_id.is_(new_id)).first()

                customer_marking = str(record_names[table][0])

                if (customer_name is not None and
                    customer_address is not None and
                    customer_marking):

                    print("\nA new record to be linked to a record:")
                    print(f"Customer's name:{customer_name[0]}")
                    print(f"Customer's address:{customer_address[0]}")
                    print(f"Marking's name:{customer_marking}")

                else:
                    print('Something went wrong!')

                    continue

                upload = (input("Proceed?\n")).lower()

                while upload != 'yes' and upload != 'no':

                    print('The answer should be only yes or no!')

                    upload = (input("Proceed?\n")).lower()

            if (new_id != 'skip' and
                new_id != 'exit' and
                new_id != 'next'):

                records_to_update.customer_id = new_id

                new_markings_df = pd.DataFrame({'marking_name':[customer_marking],
                                                'marking_customer':[customer_name[0]],
                                                'marking_customer_address':[customer_address[0]]}
                )

                print("\nA new record was added and linked:")
                print(f"Customer's name:{customer_name[0]}")
                print(f"Customer's address:{customer_address[0]}")
                print(f"Marking's name:{customer_marking}")

                if os.path.exists('C:/autocargo_reports/export/'):

                    directory_name = 'C:/autocargo_reports/export/'

                else:

                    directory_name = 'C:/'

                    print('Couldnt find a path to save the file, saving to the disc C')

                new_markings_path = f'{directory_name}new_marking.xlsx'

                new_markings_df.to_excel(new_markings_path)

                create_table(new_markings_df, Markings, session_current)

                new_markings_df.drop(0)

                session_current.execute(update_id)

                session_current.commit()

            elif new_id == 'exit':

                print('Exiting the check loop...')

                quit_for_loop = True

                break

            elif new_id == 'next':

                print(f'Skipping the {table} check loop...')

                break

            if not skip_id:

                # Fetch the records that match the empty condition
                records_to_update = session_current.query(table).filter(condition).first()

            if skip_id:

                records_to_update = session_current.query(table).filter(and_(condition, condition_list[2].not_in(skip_id))).first()

        if quit_for_loop:

            break

def empty_shipment_id(session_current):

    """
    This function fills an empty shipment_id of tables

    """

    quit_for_loop = False

    ship_id_queries = {OperationsForever:[OperationsForever.shipment_id, 'expense_forever_id'],
                       OperationsIphandlers:[OperationsIphandlers.shipment_id, 'expense_ip_id'],
                       OperationsSales:[OperationsSales.shipment_id, 'sale_id']}

    for table, condition_list in ship_id_queries.items():

        # SCondition to identify records where customer_id is None
        condition = condition_list[0].is_(None)

        # Fetch the records that match the empty condition
        records_to_update = session_current.query(table).filter(condition).all()

        for record in records_to_update:

            if table == OperationsIphandlers:

                record_names = {table:[record.expense_awb,
                                       record.expense_full_box,
                                       record.expense_marking,
                                       record.expense_eta_date,
                                       Shipments.shipment_date > record.expense_eta_date,
                                       record.expense_ip_id]}

            elif table == OperationsForever:

                if (record.expense_full_box == 0E-10 and
                    record.expense_volume == 0E-10):

                    record.shipment_id = '0'

                    continue

                record_names = {table:[record.expense_awb,
                                       record.expense_full_box,
                                       record.expense_marking,
                                       record.expense_date,
                                       Shipments.shipment_date < record.expense_date,
                                       record.expense_forever_id]}

            elif table == OperationsSales:

                record_names = {table:[record.sale_awb,
                                       record.sale_full_box,
                                       record.sale_marking,
                                       record.sale_date,
                                       Shipments.shipment_date < record.sale_date,
                                       record.sale_id]}

            new_id = ''

            upload = 'no'

            print(border_line)
            print(f"Updating record:\n{record}.")
            print("Input a shipment_id number to link, skip, next or exit")

            awb_match = session_current.query(Shipments.shipment_id,
                                              Shipments.shipment_awb,
                                              Shipments.shipment_box_full,
                                              Shipments.shipment_marking,
                                              Shipments.shipment_date).filter(and_(
                                              Shipments.shipment_awb == record_names[table][0],
                                              record_names[table][4]
                                              )
            )

            marking_match = session_current.query(Shipments.shipment_id,
                                                  Shipments.shipment_awb,
                                                  Shipments.shipment_box_full,
                                                  Shipments.shipment_marking,
                                                  Shipments.shipment_date).filter(and_(
                                                  Shipments.shipment_marking == record_names[table][2],
                                                  record_names[table][4]
                                                  )
            )

            awb_box_match = session_current.query(Shipments.shipment_id,
                                                  Shipments.shipment_awb,
                                                  Shipments.shipment_box_full,
                                                  Shipments.shipment_marking,
                                                  Shipments.shipment_date).filter(and_(
                                                  Shipments.shipment_awb == record_names[table][0],
                                                  Shipments.shipment_box_full == record_names[table][1],
                                                  record_names[table][4]
                                                  )
            )

            awb_marking_match = session_current.query(Shipments.shipment_id,
                                                      Shipments.shipment_awb,
                                                      Shipments.shipment_box_full,
                                                      Shipments.shipment_marking,
                                                      Shipments.shipment_date).filter(and_(
                                                      Shipments.shipment_awb == record_names[table][0],
                                                      Shipments.shipment_marking == record_names[table][2],
                                                      record_names[table][4]
                                                      )
            )

            box_marking_match = session_current.query(Shipments.shipment_id,
                                                      Shipments.shipment_awb,
                                                      Shipments.shipment_box_full,
                                                      Shipments.shipment_marking,
                                                      Shipments.shipment_date).filter(and_(
                                                      Shipments.shipment_box_full == record_names[table][1],
                                                      Shipments.shipment_marking == record_names[table][2],
                                                      record_names[table][4]
                                                      )
            )

            match_list = {'Found a match by AWB:':awb_match,
                          'Found a match by marking:':marking_match,
                          'Found a match by AWB and full box:':awb_box_match,
                          'Found a match by AWB and marking:':awb_marking_match,
                          'Found a match by marking and full box:':box_marking_match}

            for match_type, match in match_list.items():

                for row in match:

                    print(border_line)
                    print(match_type)
                    print(f"Shipment's ID:{row.shipment_id}")
                    print(f"Shipment's date:{row.shipment_date}")
                    print(f"Shipment's awb:{row.shipment_awb}")
                    print(f"Shipment's marking:{row.shipment_marking}")
                    print(f"Shipment's full box:{row.shipment_box_full}\n")

            while upload != 'yes':

                print(border_line)

                new_id = input("\nYour input:\n")

                if (new_id == 'skip' or
                    new_id == 'exit' or
                    new_id == 'next'):

                    print('Skipping the record...')

                    break

                picked_record = session_current.query(Shipments).filter(
                                                        Shipments.shipment_id.is_(new_id))

                for p_row in picked_record:

                    if p_row:

                        print("\nA new record to be linked to a record:")
                        print(f"Shipment's ID:{p_row.shipment_id}")
                        print(f"Shipment's date:{p_row.shipment_date}")
                        print(f"Shipment's awb:{p_row.shipment_awb}")
                        print(f"Shipment's marking:{p_row.shipment_marking}")
                        print(f"Shipment's full box:{p_row.shipment_box_full}")

                        upload = (input("Proceed?\n")).lower()

                        while upload != 'yes' and upload != 'no':

                            print('The answer should be only yes or no!')

                            upload = (input("Proceed?\n")).lower()

                    elif not p_row:
                        print('No record for this shipment_id!')

            if (new_id != 'skip' and
                new_id != 'exit' and
                new_id != 'next'):

                record.shipment_id = new_id

                new_ship_id_df = pd.DataFrame({condition_list[1]:[record_names[table][5]],
                                                'shipment_id':[new_id]}
                )

                if os.path.exists('C:/autocargo_reports/export/'):

                    directory_name = 'C:/autocargo_reports/export/'

                else:

                    directory_name = 'C:/'

                    print('Couldnt find a path to save the file, saving to the disc C')

                new_ship_id_path = f'{directory_name}new_shipment_log.xlsx'

                new_ship_id_df.to_excel(new_ship_id_path)

            elif new_id == 'exit':

                print('Exiting the check loop...')

                quit_for_loop = True

                break

            elif new_id == 'next':

                print(f'Skipping the {table} check loop...')

                break

        session_current.commit()

        if quit_for_loop:

            break
