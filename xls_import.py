#!/usr/bin/python

import os, sys
from optparse import OptionParser

try:
	import MySQLdb
except ImportError:
	print "Please install MySQLdb first."
	print "http://pypi.python.org/pypi/MySQL-python"
	exit(1)

try:
	from xlrd import open_workbook, cellname
except ImportError:
	print "Please install xlrd first."
	print "http://pypi.python.org/pypi/xlrd"
	exit(1)

class XLSImporter(object):
	"""Import data from Excel and save it to a MySQL database."""

	def __init__(self):
		self.prog = os.path.basename(sys.argv[0])

		parser = OptionParser(version="%prog 0.1", usage="%prog [options] input files")

		parser.add_option("-t", "--test", action="store_true", dest="test_mode", default=False, help="test mode: do not modify database")
		parser.add_option("-q", "--quiet", action="store_true", dest="quiet_mode", default=False, help="quiet mode")
		parser.add_option("-s", "--db-host", dest="db_host", help="MySQL hostname")
		parser.add_option("-u", "--db-user", dest="db_user", help="MySQL username")
		parser.add_option("-p", "--db-password", dest="db_password", help="MySQL password")
		parser.add_option("-d", "--db-database", dest="db_database", help="MySQL database")
		parser.add_option("-b", "--db-table", dest="db_table", help="MySQL table")
		parser.add_option("-x", "--update-if-exists", action="store", dest="update_mode", help="column name that will be used to detect and update a row in MySQL if it already exists")
		parser.add_option("-k", "--keep-if-exists", action="store", dest="keep_mode", help="column name that will be used to detect and keep a row unchanged in MySQL if it already exists")

		(self.options, self.args) = parser.parse_args()

		if self.options.update_mode and self.options.keep_mode:
			parser.error("options -x and -k are mutually exclusive")
	
		if len(self.args) == 0:
			parser.error("no input files specified")

	def message(self, m, type = None):
		if not self.options.quiet_mode:
			if type:
				print "%s: %s: %s" % (self.prog, type, m)
			else:
				print "%s: %s" % (self.prog, m)

	def valuePad(self, key):
	    return '%(' + str(key) + ')s'

	def run(self):
		self.message("connecting to database")
		self.message("host: %s" % self.options.db_host)
		self.message("user: %s" % self.options.db_user)
		self.message("password: %s" % self.options.db_password)
		self.message("database: %s" % self.options.db_database)
		self.message("table: %s" % self.options.db_table)

		# establish connection
		try:
			conn = MySQLdb.connect (host=self.options.db_host, user=self.options.db_user, passwd=self.options.db_password, db=self.options.db_database)
		except TypeError:
			self.message("cannot connect to database, check your settings", "error")
			exit(1)
		except MySQLdb.Error, e:
			self.message("cannot connect to database (%d): %s" % (e.args[0], e.args[1]), "error")
			exit(1)			
	     
		cursor = conn.cursor()

		# check that a table exists
		query = "SHOW TABLES LIKE %s"
		self.message("sql: %s" % (query % self.options.db_table))
		cursor.execute(query, self.options.db_table)
		if not cursor.rowcount:
			self.message("table '%s' does not exist" % self.options.db_table, "error")
			cursor.close()
			conn.close()
			exit(1)

		# process files
		for file in self.args:
			self.message("opening: %s" % file)
			xls = open_workbook(file)
 			sheet = xls.sheet_by_index(0)
			self.message("rows in sheet: %s" % sheet.nrows)

			columns = []
			values = []

			for row_index in range(sheet.nrows):
				
				# store column names in a list
				if row_index == 0:
					for col_index in range(sheet.ncols):
						columns.append(sheet.cell(row_index, col_index).value)
						
				# insert rows into MySQL database
				else:
					del values[:]
					for col_index in range(sheet.ncols):
						values.append(sheet.cell(row_index, col_index).value)	
					
						# check if a row already exists 
						if self.options.keep_mode == columns[col_index] or self.options.update_mode == columns[col_index]:
							query = "SELECT %s FROM %s WHERE %s = %%s" % (columns[col_index], self.options.db_table, columns[col_index])
							self.message("sql: %s" % (query % sheet.cell(row_index, col_index).value))
							cursor.execute(query, sheet.cell(row_index, col_index).value)
							needs_update = sheet.cell(row_index, col_index).value if cursor.rowcount else False

					query = ""
					row = dict(zip(columns, values))

					# update row?
					if self.options.update_mode and needs_update:
						query = "UPDATE %s SET " % self.options.db_table
						query += ', '.join(["%s = %s" % (c[1], self.valuePad(c[1])) for c in enumerate(columns)])
						query += " WHERE %s = %%(searchstring)s" % self.options.update_mode
						row['searchstring'] = needs_update
						
					# do nothing?		
					elif self.options.keep_mode and needs_update:
						self.message("keeping the row unchanged")
					
					# insert row?
					else:
						query = "INSERT INTO %s" % self.options.db_table
						query += ' ('
						query += ', '.join(row)
						query += ') VALUES ('
						query += ', '.join(map(self.valuePad, row))
						query += ')'
					
					# execute query
					if query:	
						self.message("sql: %s" % (query % row))
						if not self.options.test_mode:
								cursor.execute(query, row)
						
						
		cursor.close()
		conn.commit()
		conn.close()
		
		if not self.options.test_mode:
			self.message("success: data import completed")
		else:
			self.message("success: test completed")

if __name__ == '__main__':
	importer = XLSImporter()
	importer.run()

