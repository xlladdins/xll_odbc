// xllconnect.cpp - connect to an ODBC driver
// see - http://www.connectionstrings.com/
#include "xllodbc.h"

using namespace xll;

static AddIn xai_odbc_connect(
	Function(XLL_HANDLE, "xll_odbc_connect", "\\" CATEGORY ".CONNECT")
	.Arguments({
		Arg(XLL_CSTRING, "DSN", "is the data source name."),
		Arg(XLL_CSTRING, "_User", "is the optional user name."),
		Arg(XLL_CSTRING, "_Pass", "is the optional password."),
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return a handle to an ODBC connection using a data source name.")
	.HelpTopic("https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlconnect-function")
    .Documentation(R"(
SQLConnect establishes connections to a driver and a data source. 
The connection handle references storage of all information about the connection to the data source, 
including status, transaction state, and error information.
    )")
);
HANDLEX WINAPI xll_odbc_connect(SQLTCHAR* dsn, SQLTCHAR* user, SQLTCHAR* pass)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<ODBC::Dbc> hdbc(new ODBC::Dbc());

		ensure (SQL_SUCCEEDED(SQLConnect(*hdbc, dsn, SQL_NTS, user, SQL_NTS, pass, SQL_NTS)) || ODBC_ERROR(*hdbc));

		h = hdbc.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddIn xai_odbc_browse_connect(
	Function(XLL_HANDLE, "?xll_odbc_browse_connect", "\\ODBC.BROWSE_CONNECT")
	.Arguments({
		Arg(XLL_CSTRING, "\"DRIVER={Microsoft Excel Driver (*.xls)}\"", "is an odbc data source name."),
		})
	.Uncalced()
	.Category("ODBC")
	.FunctionHelp("Display a dialog box of available drivers and return a handle to an ODBC connection.")
);
HANDLEX WINAPI xll_odbc_browse_connect(const wchar_t* conn)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<ODBC::Dbc> ph(new ODBC::Dbc());

		while (SQL_NEED_DATA == ph->BrowseConnect(conn)) {
			// get networks
			// get servers
			// get databases
			// get tables
		}

		h = ph.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		// call ODBC::DiagField::Get for more info
	}

	return h;
}

static AddIn xai_odbc_driver_connect(
	Function(XLL_HANDLE, "xll_odbc_driver_connect", "\\ODBC.DRIVER_CONNECT")
	.Arguments({
		Arg(XLL_LPOPER, "driver", "is an odbc connection string.", "{DRIVER=\"{SQL Server}\"}"),
		})
	.Uncalced()
	.Category("ODBC")
	.FunctionHelp("Display a list of available drivers and return a handle to an ODBC connection.")
	.HelpTopic("https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqldriverconnect-function")
	.Documentation(R"(
SQLDriverConnect is an alternative to SQLConnect. 
It supports data sources that require more connection information than the three arguments in SQLConnect, 
dialog boxes to prompt the user for all connection information, 
and data sources that are not defined in the system information.
</p><p>
SQLDriverConnect provides the following connection attributes:
</p>
<ul>
<li>
<p>
    Establish a connection using a connection string that contains the data source name, one or more user IDs, one or more passwords, and other information required by the data source.
</p>
</li>
<li>
<p>
    Establish a connection using a partial connection string or no additional information; in this case, the Driver Manager and the driver can each prompt the user for connection information.
</p>
</li>
<li>
<p>
    Establish a connection to a data source that is not defined in the system information. If the application supplies a partial connection string, the driver can prompt the user for connection information.
</p>
</li>
<li>
<p>
    Establish a connection to a data source using a connection string constructed from the information in a .dsn file.
</p>
</li>
</list>
<p>
After a connection is established, SQLDriverConnect returns the completed connection string. 
The application can use this string for subsequent connection requests. 
For more information, see Connecting with SQLDriverConnect.
    )")

);
HANDLEX WINAPI xll_odbc_driver_connect(LPOPER pcs)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<ODBC::Dbc> hdbc(new ODBC::Dbc());

		const OPER& cs(*pcs);
		ensure (cs[0].xltype == xltypeStr);

		std::wstring conn(cs[0].val.str + 1, cs[0].val.str[0]);
		
		for (WORD i = 1; i < cs.size(); ++i ) {
			if (i%2 == 0)
				conn.append(1, ';');
			else
				conn.append(1, '=');

			conn.append(cs[i].val.str + 1, cs[i].val.str[0]);
		}

		ensure (SQL_SUCCEEDED(hdbc->DriverConnect((const SQLTCHAR*)conn.c_str())) || ODBC_ERROR(*hdbc));

		h = hdbc.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddIn xai_odbc_connection_string(
	Function(XLL_CSTRING, "?xll_odbc_connection_string", "ODBC.CONNECTION.STRING")
	.Arguments({
		Arg(XLL_HANDLE, "Dbc", "is a handle returned by ODBC.CONNECT.*."),
		})
	.Category("ODBC")
	.FunctionHelp("Returns the full connection string.")
    .Documentation(R"(
Return the full connection string after a successful connection to and ODBC data source.
)")
);
const SQLTCHAR* WINAPI xll_odbc_connection_string(HANDLEX dbc)
{
#pragma XLLEXPORT
	const SQLTCHAR* conn(0);

	try {
		handle<ODBC::Dbc> hdbc(dbc);

		ensure (hdbc);
		
		conn = hdbc->connectionString();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return conn;
}
