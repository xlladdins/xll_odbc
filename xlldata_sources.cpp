// xlldata_source.cpp list of available ODBC data sources
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_data_sources(
	FunctionX(XLL_LPOPERX, _T("?xll_odbc_data_sources"), _T("ODBC.DATA.SOURCES"))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Retrieve a range of data sources and descriptions."))
);
LPOPERX WINAPI xll_odbc_data_sources(void)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o.resize(0,0);
		OPERX r(1,2);
		r[0] = OPERX(_T(""), 255);
		r[1] = OPERX(_T(""), 255);

		while (SQL_NO_DATA != SQLDataSources(ODBC::Env(), SQL_FETCH_NEXT, ODBC_BUFS(r[0]), ODBC_BUFS(r[1]))) {
			o.push_back(r);
			r[0].val.str[0] = 255;
			r[1].val.str[0] = 255;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	return &o;
}
