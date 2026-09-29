ke MyODBC independent of the character set (like BIG5).
	
1999-10-01  Michael Widenius  <monty@tik.pp.sci.fi>

* Added mutex lock around all use of DBC;  This should make MyODBC thread safe
  even for multiple threads share the same connection.
* MyODBC now uses libmysqlclient from MySQL 3.23.  One can now configure
  MyODBC without having to install the MySQL source.
	
1999-09-15  Michael Widenius  <monty@monty.pp.sci.fi>

* Some small fixes in configure to help it find libmysqlclient.a

Fri Sep  3 13:34:02 1999  Michael Widenius  <monty@bitch.pp.sci.fi>

* Release of 3.22.25

1999-07-01  Michael Widenius  <monty@monty.pp.sci.fi>

* Fix for SQLSetStmtOption to return SQL_SUCCESS_WITH_INFO instead of
  SQL_ERROR if MyODBC changes the requested cursor type.
* Changed debugging to not clear the debug file on restart.
	
1999-06-27  Michael Widenius  <monty@monty.pp.sci.fi>

* Added a wrapper around SQLFreeStmt to avoid linker problems.

1999-06-07  Michael Widenius  <monty@monty.pp.sci.fi>

* Added test if ODBC debugging is requested on SQLConnect()
  (Thanks for Peter Harvey for this)	
	
1999-06-06  Michael Widenius  <monty@monty.pp.sci.fi>

* Release of 3.22.24
* Added test of SQL_BESTROW_ID in SQLSpecialColumns
	
1999-06-03  Michael Widenius  <monty@monty.pp.sci.fi>

* The MyODBC driver never reported that a column is updatable when calling
  SQLColAttributes. (Patch by rsasnett@microsoft.com)

1999-05-14  Michael Widenius  <monty@monty.pp.sci.fi>

* Fixed of 'Invalid cursor name Data truncated'
* Changed the FLAG connect option to OPTION (to be consistent with the Unix
  MyODBC version)

1999-05-01  Michael Widenius  <monty@monty.pp.sci.fi>

* Patch for big floats (by Dmitry Potapov <dpotapov@capitalsoft.com>)

1999-04-27  Michael Widenius  <monty@monty.pp.sci.fi>

* Release of 3.22.23	
	
1999-04-27  Manush Dodunekov  <manush@litecom.net>
* configure.in: Added option --with-odbc-ini=PATH
* misc.c: Modified to look for odbc.ini in the iodbc way.

1999-04-21  Manush Dodunekov  <manush@litecom.net>

* connect.c (SQLConnect): Initial statement now defaults to
  "", not "0". Connect flag is read from "option", not "flag".
  Connect-string keywords are now case insensitive on Unix.

* info.c (SQLGetInfo): Return mysql version instead of driver version
  with SQLGetInfo(SQL_DBMS_VER)

* Prepared for future libiodbcinst, provided is a temporary
  ini file reader - no patch to iodbc needed.

1999-04-20  Manush Dodunekov <manush@litecom.net>

* Ported MyODBC to IODBC.  Did also the following changes:

SWORD -> SQLSMALLINT
UWORD -> SQLUSMALLINT
SDWORD -> SQLINTEGER
UDWORD -> SQLUINTEGER
PTR -> SQLPOINTER
UCHAR -> SQLCHAR
HENV -> SQLHENV
HDBC -> SQLHDBC
HSTMT -> SQLHSTMT
	
1999-04-20  Lars Doelle <lars.doelle@on-line.de>

* Patch to make myodbc use the odbc.ini entries.

The following is the ChangeLog before the port to libiodbc:
	
2.50.23

- SQLGetStmtOption() with SQL_ROW_NUMBER is now 1 based.
- SQLColAttributes() called with SQL_COLUMN_COUNT on something that didn't
  return a result set crashed.

2.50.22a
- Fixed type error in result.c

2.50.22:
- Don't release bound columns until SQLFreeStmt with SQL_UNBIND or SQL_DROP.
- Don't give an error with MySQL 3.21 when checking if the connection is up.
- Force initialization of threads even if the application doesn't call
  LibMain() with DLL_THREAD_ATTACH.
- Added BINARY -> CHAR conversion for BLOBS.

2.50.21:
- Changed SQLGetData to not add end null to binary data that
  must be retrieved trough many SQLGetData calls.
  (The old way should have worked, but Mike Watson reported
  some problems with Access that this patch should fix)
- Added an ODBC connect option to make one SQL query at connection.
  This is useful to tell the MySQL server to use a specific character set
  when connecting.
- Upgraded the MySQL interface to 3.22.10 (To support the new options flags)

2.50.20:
- Added mapping of MySQL error codes in connect and query to
  ODBC error codes.
- Some small changes of handling CATALOGS with SQLGetInfo()
- Compressed protcol didn't work because MEMMOVE wasn't defined.
- Allow one to call SQLColAttribute with col=0 and
  type = SQL_DESC_COUNT (needed by odbcexpress).

2.50.19:
- When using SQLSetStmtOption with SQL_CURSOR_TYPE one could get
  'data truncated' errors when retrieving data.
- MyODBC 2.50.18 crashed when trying to connect from some clients.

2.50.18:
- Added option flag for extension of SQLDescribeCol
- One can now bind LONG to date columns (ODBC extension)
- MyODBC crashed when one used the following combination:
  SQLPrepare/SQLBindParameter/SQLExectute/SQLFreeSTMT(SQL_RESET_PARAM)/
  SQLBindParameter.
- Updated MySQL API to 3.22.8 with named pipe support and compression.
- Now options flags (see above)
- New nice interface to the options flags from the MyODBC connect screen.
- MyODBC now reports to MySQL that this is a ODBC client, which allows
  MySQL 3.22.7 and above to behave differently to satisfy weird requests.
- If there is more than one hour since last query, MyODBC will now check
  that the connection is ok and do a reconnect if needed.

2.50.17
- DATE fields was updated with DATE + TIME (bug from 2.50.16)
- BLOB fields are now returned as LONGVARBINARY instead of LONGVARCHAR.
  TEXT fields are still returned as LONGVARCHAR.
  (Patch by "Shayne Hughes" <shayne@vgl.ucdavis.edu>)

2.50.16
- Lots of extra parameter checking for safer code with different drivers.
- Format MySQL timestamp columns when converted to ODBC strings.
  (YYYYMMDDHHMMSS -> YYYY-MM-DD HH:MM:SS)
- Allow NULL values in auto increment columns (for Delphi 3.0)
- Added option to pad character fields with space to max length (for
  quiktest.dll)
- Fixes bind problems from 2.50.15
  NOTE: If you have problems with 2.50.16, the most stable version is 2.50.13 !
- SQLExtended_fetch() couldn't always read from overlapping start region

2.50.15
- Fixed a problem with Excel and SQLTables() from 2.50.14.
- Changed SQLBindCol() to allow one to remove old value bindings but keep
  length bindings (very seldom used, but required by the ODBC standard)
- Allow one to use SQL_CURSOR_FORWARD and SQL_CURSOR_STATIC. Fixes problem
  with ASP.
- Change locale for SQL parameters and bound float/double variables to allow
  one to have ',' as decimal point in ones locale!
  Fix also locale for all SQL parameters that are inserted as CHAR but
  requested as DOUBLE. (To fix money type in Access 97)
- If name is truncated, now return the truncated length instead of the original
  length (because of the perl DBI/ODBC driver).
- Allow one to SQLBindCol() before SQLPrepare() because of VB 5.0
  (Who came up with this BROKEN standard ?)

2.50.14
- Some small changes to SQLTable() to get some ODBC applications to work.
  Now one can request catalog names, schema names and table types from
  SQLTable().
- Added support for the new YEAR type in MySQL 3.22.
- Added support to turn of schemas in MySQL 3.22 (In MySQL 3.22 schema =
  database)
- Different releases for win95 and NT to avoid installation problems on NT.
- Added option to force use of ODBC manager cursors (See the option flag table)

2.50.13
- Added SQLGetInfo(SQL_OJ_CAPABILITIES).
- Added option flag 32
- Fixed bug with SQLGetInfo(DATABASE_NAME = 16) from MyODBC 2.50.12

2.50.12
- Change 0000-00-00 DATE values to NULL if retrieved as DATE.
- SQLParameter() extended to support all ODBC 2.50 parameter types.
- Using new connect protocol for MyODBC 3.21.26 to make faster connections.
- Don't remove \r\n from SQL statement.
- Added option flag 16

2.50.11
- SQLColumns() didn't return right column lengths().
- SQLColumns() on BLOBS columns now returns max blob length.
- SQLGetInfo() now reports that MySQL supports allmost all DATE functions.
- Fixed problem that some applications (at least PowerBuilder) uses wrong
  max column lengths for fixed types to SQLBindCol().

2.50.10
- Added support for SQLNativeSQL.
- SQLGetInfo() didn't report that MyODBC supports SQLGetConnectoption.
- Allow SQLBind() to be called before SQLExecute().

2.50.9
- VB5 does a wrong 'SELGetInfo' request where the save address is 0x.  Requests
  like these are now silenty ignored.
- SQLDescribParam() implemented by always returning VARCHAR(255). As MySQL
  can take strings anywhere this should be a workable solution!
- Changed password to allow 16 characters (was 8).

2.50.8
- More debug information to the MyODBC log.

2.50.7
- Upgraded libmysql.c to allow automatic reconnection in some cases.
- Fixed problems in SQLExtendedFetch() and SQLGetTypeInfo() that crashed
  Delphi 3.0
- DATETIME now return timestamp as default value (crashed Access)
- The distribution now contains include files and libraries for use
  with C clients.  One have to link with all included .lib files
  to get a C client.
- Use my_SQLExecute instead of SQLExecute to avoid dynamic link problems
  on some Unix systems.
- Extended SQLSetConnectOPtion and SQLGetConnectOptions to handle all
  options (some are silently ignored).
- SQLGetStmtOption now handles all ODBC 2.5 options.

2.50.6
- Report the MySQL have LEFT OUTER JOIN and all other functions from
  MySQL 3.21.17

2.50.5
- Fixed better logging by changing ASCII 0 to \0 when inserting parameters.
- Optimized 'mysql_list_tables' to not call malloc so many times.

2.50.4
- Added Option bits FLAG_BIG_PACKETS and FLAG_DEBUG
- Fixed bug with longvarbinary parameters with ASCII 0

2.50.3:
- Added Option 2 for Access, see NOTES.
- Fixed bug with date fields and Access.

2.50.2:
- Fixed DATETIME and TIME problems when using SQLBind.
- Fixed that port is saved properly.
- Added options flag for some BDE drivers. Options = 1 tells MyODBC not to
  return max_width for columns. This should help Delphi connections.
- Changed SqlColumns(), SqlSpecialColumns() and SqlColAttributes() to return
  right transfer size.
- TIMESTAMP now always returns precision 19 (because of BDE).
- Added conversion for the DISPLAY_SIZE for SQLColAttributes

2.50.1:
- The installer got linked with the ODBC 3.0 libraries and didn't work.
  Fixed by relinking with ODBC 2.50.

2.50.0:
- MyODBC has now tags for