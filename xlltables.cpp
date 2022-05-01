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
static AddIn xai_odbc_tables(
	Function(XLL_LPOPER, "xll_odbc_tables", "ODBC.TABLES")
	.Arguments({
		Arg(XLL_HANDLE, "Dbc", "is a handle to a database connection."),
		Arg(XLL_CSTRING, "Catalog", "is the optional catalog pattern. Default is \"%\"", "%"),
		Arg(XLL_CSTRING, "Schema", "is the optional schema pattern. Default is \"%\"", "%"),
		Arg(XLL_CSTRING, "Table", "is the optional table pattern. Default is \"%\"", "%"),
		Arg(XLL_CSTRING, "Type", "is the optional type list. Default is \"%\"",  "%"),
		})
	.Category("ODBC")
	.FunctionHelp("Return table information.")
    .Documentation(R"(
SQLTables returns the list of table, catalog, or schema names, 
and table types, stored in a specific data source. 
The driver returns the information as a result set.
)")
);
LPOPER WINAPI xll_odbc_tables(HANDLEX h, SQLTCHAR* cat, SQLTCHAR* schem, SQLTCHAR* name, SQLTCHAR* type)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o.resize(0,0);
		handle<ODBC::Dbc> dbc(h);
		ODBC::Stmt stmt(*dbc);

//		if (!(*cat||*schem||*name||*type))
//			cat = _T("%");

		ensure (SQL_SUCCEEDED(SQLTables(stmt, cat, SQL_NTS, schem, SQL_NTS, name, SQL_NTS, type, SQL_NTS)) || ODBC_ERROR(stmt));

		OPER row(1, 5);
		for (WORD i = 0; i < 5; ++i) {
			row[i] = OPER("", 255);
			SQLLEN r0 = 254;
			ensure (SQL_SUCCEEDED(SQLBindCol(stmt, i + 1, SQL_C_CHAR, ODBC_STR(row[i]), &r0)) || ODBC_ERROR(stmt));
			row[i].val.str[0] = (SQLTCHAR)r0;
		}

		while (SQL_SUCCEEDED(SQLFetch(stmt)) || ODBC_ERROR(stmt)) {
			o.push_back(row);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(ErrNA);
	}

	if (o == OPER())
		o = OPER(ErrNull);

	return &o;
}

