// xlldata_source.cpp list of available ODBC data sources
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_data_sources(
	FunctionX(XLL_LPOPER, L"?xll_odbc_data_sources", L"ODBC.DATA.SOURCES")
	.Category(L"ODBC")
	.FunctionHelp(L"Retrieve a range of data sources and descriptions.")
    .Documentation(LR"(
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
		r[0] = OPER(L"", 255);
		r[1] = OPER(L"", 255);

		while (SQL_NO_DATA != SQLDataSources(ODBC::Env(), SQL_FETCH_NEXT, ODBC_BUFS(r[0]), ODBC_BUFS(r[1]))) {
			o.push_back(r);
			r[0].val.str[0] = 255;
			r[1].val.str[0] = 255;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(xlerr::NA);
	}

	return &o;
}
