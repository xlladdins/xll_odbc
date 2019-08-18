// xllconnect.cpp - connect to an ODBC driver
// see - http://www.connectionstrings.com/
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_connect"), _T("ODBC.CONNECT"))
	.Arg(XLL_CSTRINGX, _T("DSN"), _T("is the data source name."))
	.Arg(XLL_CSTRINGX, _T("_User"), _T("is the optional user name."))
	.Arg(XLL_CSTRINGX, _T("_Pass"), _T("is the optional password."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Return a handle to an ODBC connection."))
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
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_browse_connect"), _T("ODBC.CONNECT.BROWSE"))
	.Arg(XLL_CSTRINGX, _T("\"DRIVER={Microsoft Excel Driver (*.xls)}\""), _T("is an odbc data source name."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Display a list of available drivers and return a handle to an ODBC connection."))
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
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_driver_connect"), _T("ODBC.CONNECT.DRIVER"))
	.Range(_T("{DRIVER=\"{SQL Server}\"}"), _T("is an odbc connection string."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Display a list of available drivers and return a handle to an ODBC connection."))
);
HANDLEX WINAPI xll_odbc_driver_connect(LPOPERX pcs)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<ODBC::Dbc> hdbc(new ODBC::Dbc());

		const OPERX& cs(*pcs);
		ensure (cs[0].xltype == xltypeStr);

		xstring conn(cs[0].val.str + 1, cs[0].val.str[0]);
		
		for (xword i = 1; i < cs.size(); ++i ) {
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
	FunctionX(XLL_CSTRINGX, _T("?xll_odbc_connection_string"), _T("ODBC.CONNECTION.STRING"))
	.Handle(_T("Dbc"), _T("is a handle returned by ODBC.CONNECT.*."))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Returns the full connection string."))
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
