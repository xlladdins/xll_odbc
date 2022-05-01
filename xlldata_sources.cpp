// xlldata_source.cpp list of available ODBC data sources
#include "xllodbc.h"

using namespace xll;

static AddIn xai_odbc_data_sources(
	Function(XLL_LPOPER, "xll_odbc_data_sources", "ODBC.DATA.SOURCES")
	.Category("ODBC")
	.FunctionHelp("Retrieve a range of data sources and descriptions.")
    .Documentation(R"(
SQLDataSources returns information about a data source. This function is implemented only by the Driver Manager.
The driver determines how data source names are mapped to actual data sources.
)")
);
LPOPER WINAPI xll_odbc_data_sources(void)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o.resize(0,0);
		OPER r(1,2);
		r[0] = OPER("", 255);
		r[1] = OPER("", 255);

		SQLSMALLINT r0, r1;
		while (SQL_NO_DATA != SQLDataSources(ODBC::Env(), SQL_FETCH_NEXT, ODBC_STR(r[0]), &r0, ODBC_STR(r[1]), &r1)) {
			r[0].val.str[0] = r0;
			r[1].val.str[0] = r1;
			o.push_back(r);
			r[0].val.str[0] = 255;
			r[1].val.str[0] = 255;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(ErrNA);
	}

	return &o;
}
