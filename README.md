xlsimport
=========

xlsimport is a command line tool that imports data from XLS spreadsheets to MySQL.

What it does?
-------------

```
./xlsimport --help

Usage: xlsimport [options] input files

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -t, --test            test mode: do not modify database
  -q, --quiet           quiet mode
  -s DB_HOST, --db-host=DB_HOST
                        MySQL hostname
  -u DB_USER, --db-user=DB_USER
                        MySQL username
  -p DB_PASSWORD, --db-password=DB_PASSWORD
                        MySQL password
  -d DB_DATABASE, --db-database=DB_DATABASE
                        MySQL database
  -b DB_TABLE, --db-table=DB_TABLE
                        MySQL table
  -x UPDATE_MODE, --update-if-exists=UPDATE_MODE
                        column name that will be used to detect and update a
                        row in MySQL if it already exists
  -k KEEP_MODE, --keep-if-exists=KEEP_MODE
                        column name that will be used to detect and keep a row
                        unchanged in MySQL if it already exists
```

How it works?
-------------

```
mysql --database=xlsimport --user=root < test.sql
./xlsimport --db-host=localhost --db-user=root --db-password='' --db-database=xlsimport --db-table=test test.xls
xlsimport: connecting to database
xlsimport: host: localhost
xlsimport: user: root
xlsimport: password: 
xlsimport: database: xls_import
xlsimport: table: test
xlsimport: sql: SHOW TABLES LIKE test
xlsimport: opening: test.xls
xlsimport: rows in sheet: 6
xlsimport: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (blue, green, black)
xlsimport: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (yellow, black, pink)
xlsimport: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (violet, green, yellow)
xlsimport: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (brown, blue, green)
xlsimport: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (white, pink, blue)
xlsimport: success: data import completed
```

Dependencies
------------

You will need Python 2.7, [MySQLdb](http://pypi.python.org/pypi/MySQL-python) and [xlrd](http://pypi.python.org/pypi/xlrd) installed to use this tool.
