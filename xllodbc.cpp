// xllodbc.cpp - ODBC wrappers for Excel
#include "xllodbc.h"

using namespace xll;
/*
static AddIn xai_odbc_documentation(
    Documentation(R"(
This add-in provides an ODBC interface for Excel.
)"));
*/

static AddIn xai_odbc_get_info(
	Function(XLL_LPOPER, "xll_odbc_get_info", "ODBC.GET.INFO")
	.Arguments({
		Arg(XLL_HANDLE, "Dbc", "is a handle to a database connection."),
		Arg(XLL_USHORT, "Type", "is a type from the ODBC_INFO_TYPE_* enumeration."),
		})
	.Category("ODBC")
	.FunctionHelp("Returns general information about the driver and data source associated with a connection.")
    .Documentation(R"(
SQLGetInfo returns general information about the driver and data source associated with a connection.
)")
);
LPOPER WINAPI xll_odbc_get_info(HANDLEX dbc, USHORT type)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<ODBC::Dbc> hdbc(dbc);

		o = OPER("", 255);
		SQLSMALLINT o0;
	    ensure (SQL_SUCCESS == SQLGetInfo(*hdbc, type, ODBC_STR(o), &o0));
		o.val.str[0] = o0;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(ErrNA);
	}

	return &o;
}
#if 0
static AddIn xai_odbc_execute(
	Function(XLL_LPOPER, "xll_odbc_execute", "ODBC.EXECUTE")
	.Arguments({
		Arg(XLL_HANDLE, "Dbc", "is a handle to a database connection."),
		Arg(XLL_LPOPER, "Query", "is a SQL query."),
		})
	.Category("ODBC")
	.FunctionHelp("Return a handle to the result of a query.")
    .Documentation(R"(
Prepare, execute, and fetch the results of Query.
)")
);
LPOPER WINAPI xll_odbc_execute(HANDLEX dbc, LPOPER pq)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<ODBC::Dbc> hdbc(dbc);
		ODBC::Stmt stmt(*hdbc);

		const OPER& q(*pq);
		for (WORD i = 0; i < q.size(); ++i) {
			ensure (SQL_SUCCEEDED(SQLPrepare(stmt, ODBC_STR(q[i]))) || ODBC_ERROR(stmt));
		}

//		ensure (SQL_SUCCEEDED(SQLExecute(stmt)));
		ensure (SQL_SUCCEEDED(SQLExecute(stmt)) || ODBC_ERROR(stmt));

		o.resize(0,0);
		ODBC::Bind row(stmt);
//		double x;
//		SQLBindCol(stmt, 2, SQL_C_DOUBLE, &x, sizeof(double), 0);
		while (SQL_SUCCEEDED(SQLFetch(stmt)) || ODBC_ERROR(stmt)) {
			row.getData();
			o.push_back(row);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		
		o = OPER(ErrNA);
	}

	return &o;
}

#ifdef _DEBUG

void xll_test_connect(void)
{
	ODBC::Dbc dbc;
	dbc.BrowseConnect((const SQLTCHAR*)"DSN=foo");
	
}

int xll_test_odbc()
{
	try {
		xll_test_connect();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		 
		return 0;
	}

	return 1;
}
static Auto<OpenAfter> xao_test_odbc(xll_test_odbc);

#endif
#endif // 0