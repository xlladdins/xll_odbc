// xllodbc.h
//#define EXCEL12
#include "xll/xll/xll.h"
#include "odbc.h"

template<enum ODBC::SQL_HANDLE T>
inline bool ODBC_ERROR(ODBC::Handle<T>& h)
{
	auto msg = h.GetDiagRec();
	if (msg.length()) {
		MessageBox(0, h.GetDiagRec().c_str(), L"Error", MB_OK);
		return true;
	}

	return false;
}

namespace xll {

	inline OPER GetData(ODBC::Stmt& stmt, SQLUSMALLINT n, SQLSMALLINT type, SQLULEN size)
	{
		OPER o;
		SQLRETURN ret = SQL_ERROR;

		switch (type) {
		case SQL_TYPE_DATE:
		case SQL_TYPE_TIME:
		case SQL_TYPE_TIMESTAMP:
			TIMESTAMP_STRUCT ts;
			ret = SQLGetData(stmt, n + 1, type, &ts, sizeof(ts), nullptr);
			if (SQL_SUCCEEDED(ret)) {
				double d = 0;
				if (ts.year)
					d += xll::Excel(xlfDate, xll::OPER(ts.year), xll::OPER(ts.month), xll::OPER(ts.day)).as_num();
				if (ts.hour || ts.minute || ts.second)
					d += xll::Excel(xlfTime, xll::OPER(ts.hour), xll::OPER(ts.minute), xll::OPER(ts.second)).as_num();
				o = d;
			}
			break;
		case SQL_CHAR:
		case SQL_VARCHAR:
			{
				SQLTCHAR len = static_cast<SQLTCHAR>(size);
				o = OPER("", len + 1);
				o.val.str[0] = len;
				ret = SQLGetData(stmt, n + 1, type, ODBC_BUF_(SQLLEN, o));
			}
			break;
		case SQL_INTEGER:
			SQLINTEGER i;
			ret = SQLGetData(stmt, n + 1, type, &i, 0, nullptr);
			o = i;
			break;
		case SQL_SMALLINT:
			SQLSMALLINT si;
			ret = SQLGetData(stmt, n + 1, type, &si, 0, nullptr);
			o = si;
			break;
		case SQL_FLOAT:
		case SQL_DOUBLE:
			SQLFLOAT f;
			ret = SQLGetData(stmt, n + 1, type, &f, 0, nullptr);
			o = f;
			break;
		case SQL_REAL:
			SQLREAL r;
			ret = SQLGetData(stmt, n + 1, type, &r, 0, nullptr);
			o = r;
			break;
		}

		if (!SQL_SUCCEEDED(ret)) {
			o = ErrValue;
		}

		return o;
	}
}
