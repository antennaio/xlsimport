xls_import.py
=============

xls_import.py is a command line tool that imports data from XLS spreadsheets to MySQL.

What it does?
-------------

```
./xls_import.py --help

Usage: xls_import.py [options] input files

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
mysql --database=xls_import --user=root < test.sql
./xls_import.py --db-host=localhost --db-user=root --db-password='' --db-database=xls_import --db-table=test test.xls
xls_import.py: connecting to database
xls_import.py: host: localhost
xls_import.py: user: root
xls_import.py: password: 
xls_import.py: database: xls_import
xls_import.py: table: test
xls_import.py: sql: SHOW TABLES LIKE test
xls_import.py: opening: test.xls
xls_import.py: rows in sheet: 6
xls_import.py: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (blue, green, black)
xls_import.py: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (yellow, black, pink)
xls_import.py: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (violet, green, yellow)
xls_import.py: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (brown, blue, green)
xls_import.py: sql: INSERT INTO test (value_a, value_x, value_b) VALUES (white, pink, blue)
xls_import.py: success: data import completed
```

Dependencies
------------

You will need [MySQLdb](http://pypi.python.org/pypi/MySQL-python) and [xlrd](http://pypi.python.org/pypi/xlrd) installed to use this tool.
