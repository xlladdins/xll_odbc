// xllodbc.cpp - ODBC wrappers for Excel
#include "xllodbc.h"

using namespace xll;

static AddIn xai_odbc_documentation(
    Documentation(LR"(
This add-in provides an ODBC interface for Excel.
)"));

static AddIn xai_odbc_get_info(
	Function(XLL_LPOPER, L"?xll_odbc_get_info", L"ODBC.GET.INFO")
	.Arg(XLL_HANDLE, L"Dbc", L"is a handle to a database connection.")
	.Arg(XLL_USHORT, L"Type", L"is a type from the ODBC_INFO_TYPE_* enumeration.")
	.Category(L"ODBC")
	.FunctionHelp(L"Returns general information about the driver and data source associated with a connection.")
    .Documentation(LR"(
SQLGetInfo returns general information about the driver and data source associated with a connection.
</para><para>

)")
);
LPOPER WINAPI xll_odbc_get_info(HANDLEX dbc, USHORT type)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<ODBC::Dbc> hdbc(dbc);

	    ensure (SQL_SUCCESS == SQLGetInfo(*hdbc, type, ODBC_BUFS(o)));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(xlerr::NA);
	}

	return &o;
}

static AddIn xai_odbc_execute(
	Function(XLL_LPOPER, L"?xll_odbc_execute", L"ODBC.EXECUTE")
	.Arg(XLL_HANDLE, L"Dbc", L"is a handle to a database connection.")
	.Arg(XLL_LPOPER, L"Query", L"is a SQL query.")
	.Category(L"ODBC")
	.FunctionHelp(L"Return a handle to the result of a query.")
    .Documentation(LR"(
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
		
		o = OPER(xlerr::NA);
	}

	return &o;
}

#ifdef _DEBUG

void xll_test_connect(void)
{
	ODBC::Dbc dbc;
	dbc.BrowseConnect((const SQLTCHAR*)L"DSN=foo");
	
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