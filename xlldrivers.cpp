// xlldrivers.cpp - list of ODBC drivers
#include "xllodbc.h"

using namespace xll;

// "parse\0null\0terminated\0sequence\0\0
//inline OPER split0(xcstr s)
//{
//	OPER o;
//
//	for (xcstr b = s, e = wcschr(s, 0); e[1]; b = e + 1, e = wcschr(b, 0)) {
//		o.push_back(OPER(b, e - b));
//	}
//
//	return o;
//}

static AddIn xai_odbc_drivers(
	Function(XLL_LPOPER, "xll_odbc_drivers", "ODBC.DRIVERS")
	.Uncalced()
	.Category("ODBC")
	.FunctionHelp("Retrieve a range of ODBC drivers.")
    .Documentation(R"(
SQLDrivers lists driver descriptions and driver attribute keywords. 
This function is implemented only by the Driver Manager.
</p><p>
SQLDrivers returns the driver description in the *DriverDescription buffer. 
It returns additional information about the driver in the *DriverAttributes buffer as a list of 
keyword-value pairs. 
All keywords listed in the system information for drivers will be returned for all drivers, 
except for CreateDSN, which is used to prompt creation of data sources and therefore is optional.
)")
);
LPOPER WINAPI xll_odbc_drivers(void)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o.resize(0,0);
		OPER r(1,2);
		r[0] = OPER("", 255);
		r[1] = OPER("", 255);

		SQLSMALLINT r0 = 254, r1 = 254;
		while (SQL_NO_DATA != SQLDrivers(ODBC::Env(), SQL_FETCH_NEXT, ODBC_STR(r[0]), &r0, ODBC_STR(r[1]), &r1)) {
			r[0].val.str[0] = r0;
			r[1].val.str[0] = r1;
			wchar_t* pr(r[1].val.str);
			// "char\0char\0\0" -> "char;char;\0"
			for (WORD i = 1; i <= pr[0]; ++i) {
				if (!pr[i] && pr[i+1])
					pr[i] = ';';
			}
			o.push_back(r);

			r[0].val.str[0] = 254;
			r[1].val.str[0] = 254;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(ErrNA);
	}

	return &o;
}

