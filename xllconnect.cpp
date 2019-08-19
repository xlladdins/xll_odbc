// xllconnect.cpp - connect to an ODBC driver
// see - http://www.connectionstrings.com/
#include "xllodbc.h"

using namespace xll;

static AddIn xai_odbc_connect(
	Function(XLL_HANDLE, L"?xll_odbc_connect", L"ODBC.CONNECT")
	.Arg(XLL_CSTRING, L"DSN", L"is the data source name.")
	.Arg(XLL_CSTRING, L"?User", L"is the optional user name.")
	.Arg(XLL_CSTRING, L"?Pass", L"is the optional password.")
	.Uncalced()
	.Category(L"ODBC")
	.FunctionHelp(L"Return a handle to an ODBC connection.")
    .Documentation(LR"(
SQLConnect establishes connections to a driver and a data source. 
The connection handle references storage of all information about the connection to the data source, 
including status, transaction state, and error information.
    )")
);
HANDLEX WINAPI xll_odbc_connect(SQLTCHAR* dsn, SQLTCHAR* user, SQLTCHAR* pass)
{
#pragma XLLEXPORT
	handlex h;

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
/* not used
static AddInX xai_odbc_browse_connect(
	FunctionX(XLL_HANDLEX, L"?xll_odbc_browse_connect"), L"ODBC.CONNECT.BROWSE"))
	.Arg(XLL_CSTRINGX, L"\"DRIVER={Microsoft Excel Driver (*.xls)}\""), L"is an odbc data source name."))
	.Uncalced()
	.Category(L"ODBC"))
	.FunctionHelp(L"Display a list of available drivers and return a handle to an ODBC connection."))
);
HANDLEX WINAPI xll_odbc_browse_connect(sqlcstr conn)
{
#pragma XLLEXPORT
	handlex h;

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
*/
static AddInX xai_odbc_driver_connect(
	FunctionX(XLL_HANDLE, L"?xll_odbc_driver_connect", L"ODBC.CONNECT.DRIVER")
	.Arg(XLL_LPOPER, L"driver", L"is an odbc connection string.", L"{DRIVER=\"{SQL Server}\"}")
	.Uncalced()
	.Category(L"ODBC")
	.FunctionHelp(L"Display a list of available drivers and return a handle to an ODBC connection.")
    .Documentation(LR"(
SQLDriverConnect is an alternative to SQLConnect. 
It supports data sources that require more connection information than the three arguments in SQLConnect, 
dialog boxes to prompt the user for all connection information, 
and data sources that are not defined in the system information.
</para><para>
SQLDriverConnect provides the following connection attributes:
</para>
<list class="bullet">
<listItem>
<para>
    Establish a connection using a connection string that contains the data source name, one or more user IDs, one or more passwords, and other information required by the data source.
</para>
</listItem>
<listItem>
<para>
    Establish a connection using a partial connection string or no additional information; in this case, the Driver Manager and the driver can each prompt the user for connection information.
</para>
</listItem>
<listItem>
<para>
    Establish a connection to a data source that is not defined in the system information. If the application supplies a partial connection string, the driver can prompt the user for connection information.
</para>
</listItem>
<listItem>
<para>
    Establish a connection to a data source using a connection string constructed from the information in a .dsn file.
</para>
</listItem>
</list>
<para>
After a connection is established, SQLDriverConnect returns the completed connection string. 
The application can use this string for subsequent connection requests. 
For more information, see Connecting with SQLDriverConnect.
    )")

);
HANDLEX WINAPI xll_odbc_driver_connect(LPOPER pcs)
{
#pragma XLLEXPORT
	handlex h;

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

static AddInX xai_odbc_connection_string(
	FunctionX(XLL_CSTRING, L"?xll_odbc_connection_string", L"ODBC.CONNECTION.STRING")
	.Arg(XLL_HANDLE, L"Dbc", L"is a handle returned by ODBC.CONNECT.*.")
	.Category(L"ODBC")
	.FunctionHelp(L"Returns the full connection string.")
    .Documentation(LR"(
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
