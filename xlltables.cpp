// xlltables.cpp - Query availabe tables on a connection.
#include "xllodbc.h"

using namespace xll;

/* Put in documentaton!!!


Column name
 

Column number
 

Data type
 

Comments
 

TABLE_CAT (ODBC 1.0)
 
1
 
Varchar
 
Catalog name; NULL if not applicable to the data source. If a driver supports catalogs for some tables but not for others, such as when the driver retrieves data from different DBMSs, it returns an empty string ("") for those tables that do not have catalogs.
 

TABLE_SCHEM (ODBC 1.0)
 
2
 
Varchar
 
Schema name; NULL if not applicable to the data source. If a driver supports schemas for some tables but not for others, such as when the driver retrieves data from different DBMSs, it returns an empty string ("") for those tables that do not have schemas.
 

TABLE_NAME (ODBC 1.0)
 
3
 
Varchar
 
Table name.
 

TABLE_TYPE (ODBC 1.0)
 
4
 
Varchar
 
Table type name; one of the following: "TABLE", "VIEW", "SYSTEM TABLE", "GLOBAL TEMPORARY", "LOCAL TEMPORARY", "ALIAS", "SYNONYM", or a data source–specific type name.

The meanings of "ALIAS" and "SYNONYM" are driver-specific.
 

REMARKS (ODBC 1.0)
 
5
 
Varchar
 
A description of the table.
*/ 
static AddInX xai_odbc_tables(
	FunctionX(XLL_LPOPERX, _T("?xll_odbc_tables"), _T("ODBC.TABLES"))
	.Handle(_T("Dbc"), _T("is a handle to a database connection."))
	.Arg(XLL_CSTRINGX, _T("Catalog"), _T("is the optional catalog pattern. Default is \"%\""), _T("%")) 
	.Arg(XLL_CSTRINGX, _T("Schema"), _T("is the optional schema pattern. Default is \"%\""), _T("%")) 
	.Arg(XLL_CSTRINGX, _T("Table"), _T("is the optional table pattern. Default is \"%\""), _T("%")) 
	.Arg(XLL_CSTRINGX, _T("Type"), _T("is the optional type list. Default is \"%\""),  _T("%"))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Return table information"))
);
LPXLOPERX WINAPI xll_odbc_tables(HANDLEX h, SQLTCHAR* cat, SQLTCHAR* schem, SQLTCHAR* name, SQLTCHAR* type)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o.resize(0,0);
		handle<ODBC::Dbc> dbc(h);
		ODBC::Stmt stmt(*dbc);

//		if (!(*cat||*schem||*name||*type))
//			cat = _T("%");

		ensure (SQL_SUCCEEDED(SQLTables(stmt, cat, SQL_NTS, schem, SQL_NTS, name, SQL_NTS, type, SQL_NTS)) || ODBC_ERROR(stmt));

		OPERX row(1, 5);
		for (xword i = 0; i < 5; ++i) {
			row[i] = OPERX(_T(""), 255);
			ensure (SQL_SUCCEEDED(SQLBindCol(stmt, i + 1, SQL_C_CHAR, ODBC_BUFI(row[i]))) || ODBC_ERROR(stmt));
		}

		while (SQL_SUCCEEDED(SQLFetch(stmt)) || ODBC_ERROR(stmt)) {
			o.push_back(row);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	if (o == OPERX())
		o = OPERX(xlerr::Null);

	return &o;
}

