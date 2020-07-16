# Amazon Accounting Report

## Description

The firm calculates sales from Amazon market place based on VAT (value added tax) region and currency segments.

Program takes an export text file from Amazon and prepares a flexible Accounting Report in `xlsx` format.

Project is heavily based on [previous work](https://github.com/yomajo/Amazon-Orders-Parser).

### Project Caveats:

* Not uploading source files, or databases due to sensitivity of personal information;
* No Excel side (VBA) implementation is uploaded;

## Features

* Filters out:
    * today's orders (assumes incomplete date);
    * orders alreadt processed before (present in database)
* Logs, backups database;
* Automatic database self-flushing of records as defined by `ORDERS_ARCHIVE_DAYS` in [orders_db.py](https://github.com/yomajo/Amazon-Accounting-Report/blob/master/Helper%20Files/orders_db.py);
* Creates a Excel report with:
    * Datasheets for each present segments in loaded raw text file with selected data for each order;
    * Summary sheet

## Example Report Screenshots

Example of report summary sheet:

![Report Summary Sceenshot](https://user-images.githubusercontent.com/45366313/87704500-286d3b80-c7a5-11ea-9877-2e83342dba0c.png)

Example of report data sheet for specific segment:

![Report Datasheet Sample](https://user-images.githubusercontent.com/45366313/85259048-0667ee00-b471-11ea-8772-b02cb091377a.jpg)



## Requirements

**Python 3.7+** 

Most requirements are for compiling python executable for Windows. `openpyxl` is the only third-party library used.

``pip install requirements.txt``