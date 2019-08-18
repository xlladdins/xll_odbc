// xllodbc.cpp - ODBC wrappers for Excel
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_get_info(
	FunctionX(XLL_LPOPERX, _T("?xll_odbc_get_info"), _T("ODBC.GET.INFO"))
	.Handle(_T("Dbc"), _T("is a handle to a database connection."))
	.Arg(XLL_USHORTX, _T("Type"), _T("is a type from the ODBC_INFO_TYPE_* enumeration."))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Returns general information about the driver and data source associated with a connection."))
);
LPOPERX WINAPI xll_odbc_get_info(HANDLEX dbc, USHORT type)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		handle<ODBC::Dbc> hdbc(dbc);

		SQLRETURN rc = SQLGetInfo(*hdbc, type, ODBC_BUFS(o));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	return &o;
}

static AddInX xai_odbc_execute(
	FunctionX(XLL_LPOPERX, _T("?xll_odbc_execute"), _T("ODBC.EXECUTE"))
	.Handle(_T("Dbc"), _T("is a handle to a database connection."))
	.Range(_T("Query"), _T("is a SQL query."))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Return a handle to the result of a query."))
);
LPOPERX WINAPI xll_odbc_execute(HANDLEX dbc, LPOPERX pq)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		handle<ODBC::Dbc> hdbc(dbc);
		ODBC::Stmt stmt(*hdbc);

		const OPERX& q(*pq);
		for (xword i = 0; i < q.size(); ++i) {
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
		
		o = OPERX(xlerr::NA);
	}

	return &o;
}

#ifdef _DEBUG

void xll_test_connect(void)
{
	ODBC::Dbc dbc;
	dbc.BrowseConnect((const SQLTCHAR*)_T("DSN=foo"));
	
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
static Auto<OpenAfterX> xao_test_odbc(xll_test_odbc);

#endif