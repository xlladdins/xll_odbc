// xlldrivers.cpp - list of ODBC drivers
#include "xllodbc.h"

using namespace xll;

// "parse\0null\0terminated\0sequence\0\0
inline OPERX split0(xcstr s)
{
	OPERX o;

	for (xcstr b = s, e = _tcschr(s, 0); e[1]; b = e + 1, e = _tcschr(b, 0)) {
		o.push_back(OPERX(b, e - b));
	}

	return o;
}

static AddInX xai_odbc_drivers(
	FunctionX(XLL_LPOPERX, _T("?xll_odbc_drivers"), _T("ODBC.DRIVERS"))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Retrieve a range of ODBC drivers."))
);
LPOPERX WINAPI xll_odbc_drivers(void)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o.resize(0,0);
		OPERX r(1,2);
		r[0] = OPERX(_T(""), 255);
		r[1] = OPERX(_T(""), 255);

		while (SQL_NO_DATA != SQLDrivers(ODBC::Env(), SQL_FETCH_NEXT, ODBC_BUFS(r[0]), ODBC_BUFS(r[1]))) {
			xchar* r1(r[1].val.str);
			// "char\0char\0\0" -> "char;char;\0"
			for (xword i = 1; i <= r1[0]; ++i) {
				if (!r1[i] && r1[i+1])
					r1[i] = ';';
			}
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

