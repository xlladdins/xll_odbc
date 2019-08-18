// xllconnect.cpp - connect to an ODBC driver
// see - http://www.connectionstrings.com/
#include "xllodbc.h"

using namespace xll;

static AddIn xai_odbc_connect(
	Function(XLL_HANDLE, L"?xll_odbc_connect", L"ODBC.CONNECT")
	.Arg(XLL_CSTRING, L"DSN", L"is the data source name.")
	.Arg(XLL_CSTRING, L"_User", L"is the optional user name.")
	.Arg(XLL_CSTRING, L"_Pass", L"is the optional password.")
	.Uncalced()
	.Category(L"ODBC")
	.FunctionHelp(L"Return a handle to an ODBC connection.")
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
	.Arg(XLL_LPOPER, L"{DRIVER=\"{SQL Server}\"}", L"is an odbc connection string.")
	.Uncalced()
	.Category(L"ODBC")
	.FunctionHelp(L"Display a list of available drivers and return a handle to an ODBC connection.")
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
