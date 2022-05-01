// xllodbc.h
//#define EXCEL12
#include "xll/xll/xll.h"
#include "odbc.h"

#define ODBC_STR(o) reinterpret_cast<SQLTCHAR*>(o.val.str + 1), o.val.str[0]

template<enum ODBC::SQL_HANDLE T>
inline bool ODBC_ERROR(ODBC::Handle<T>& h)
{
	ODBC::DiagRec dr;
	std::basic_string<SQLTCHAR> rec;

	for (SQLSMALLINT i = 1; h.Get(i, dr) != SQL_NO_DATA; ++i) {
		rec.append(dr.state);
		rec.append({ ':', ' ' });
		rec.append(dr.message);
		//rec.append('\n');
	}

	if (rec.size())
		XLL_ERROR((char*)rec.c_str());

	return false;
}
#if 0
namespace ODBC {

	template<class T>
	class lenptr {
		xll::OPER& o_;
		T len;
	public:
		lenptr(xll::OPER& o)
			: o_(o), len(0)
		{ }
		~lenptr()
		{
			if (len)
				o_.val.str[0] = static_cast<wchar_t>(len);
		}
		operator T*()
		{
			return &len;
		}
	};
#endif // 0
	inline SQLRETURN GetNum(ODBC::Stmt& stmt, SQLUSMALLINT n, xll::OPER& o)
	{
		o.xltype = xltypeNum;

		return  SQLGetData(stmt, n + 1, SQL_C_DOUBLE, &o.val.num, sizeof(double), 0);
	}
	inline SQLRETURN GetStr(ODBC::Stmt& stmt, SQLUSMALLINT n, xll::OPER& o, SQLLEN len = 0)
	{
		if (len != 0)
			o = xll::OPER("", static_cast<wchar_t>(len + 1));

		return SQLGetData(stmt, n + 1, SQL_C_TCHAR, ODBC_STR(o), &len);
	}
	inline SQLRETURN GetDate(ODBC::Stmt& stmt, SQLUSMALLINT n, xll::OPER& o)
	{
		SQLRETURN rc(SQL_ERROR);
		SQL_TIMESTAMP_STRUCT ts;

		rc = SQLGetData(stmt, n + 1, SQL_C_TYPE_TIMESTAMP, &ts, sizeof(ts), 0);
		if (SQL_SUCCEEDED(rc)) {
			o = 0;
			if (ts.year)
				o &= xll::Excel(xlfDate, xll::OPER(ts.year), xll::OPER(ts.month), xll::OPER(ts.day));
			if (ts.hour || ts.minute || ts.second)
				o &= xll::Excel(xlfTime, xll::OPER(ts.hour), xll::OPER(ts.minute), xll::OPER(ts.second));
		}

		return rc;
	}
#if 0
	struct Bind : public xll::OPER {
		xll::OPER header;
		std::vector<SQLSMALLINT> type, nullable, digits;
		std::vector<SQLULEN> len;
		std::vector<std::function<SQLRETURN(SQLSMALLINT)>> GetData;
		ODBC::Stmt& stmt_;
		Bind(ODBC::Stmt& stmt)
			: xll::OPER(1, NumResultsCols(stmt)), header(1,size()), stmt_(stmt), GetData(size()), 
				type(size()), nullable(size()), digits(size()), len(size())
		{
			for (WORD i = 0; i < size(); ++i) {
				header[i] = xll::OPER("", 255);

				ensure (SQL_SUCCEEDED(SQLDescribeCol(stmt, i + 1, ODBC_BUFS(header[i]), &type[i], &len[i], &digits[i], &nullable[i]))
					|| ODBC_ERROR(stmt));

				switch (type[i]) {
				case SQL_CHAR: case SQL_VARCHAR: case SQL_WCHAR: case SQL_WVARCHAR: 
					//!!! check for long strings!!!
					operator[](i) = xll::OPER("", static_cast<wchar_t>(len[i]));
					GetData[i] = [this](SQLSMALLINT i) { return this->GetStr(i); };

					break;

				case SQL_SMALLINT: case SQL_INTEGER: case SQL_DOUBLE: case SQL_REAL:
				case SQL_BIT: case SQL_TINYINT: case SQL_BIGINT: case SQL_FLOAT:
					GetData[i] = [this](SQLSMALLINT i) { return this->GetNum(i); };
					
					break; 

				case SQL_DATE: case SQL_TIME: case SQL_TIMESTAMP:
				case SQL_TYPE_DATE: case SQL_TYPE_TIME: case SQL_TYPE_TIMESTAMP:
					GetData[i] = [this](SQLSMALLINT i) { return this->GetDate(i); };
					
					break;

					//				default:
//					throw std::runtime_error("ODBC::Bind: unrecognized data type");
				}
			}
		}
		~Bind()
		{ }

		SQLRETURN GetNum(SQLUSMALLINT i)
		{
			return ODBC::GetNum(stmt_, i, operator[](i));
		}
		SQLRETURN GetStr(SQLUSMALLINT i)
		{
			return ODBC::GetStr(stmt_, i, operator[](i));
		}
		SQLRETURN GetDate(SQLUSMALLINT i)
		{
			return ODBC::GetDate(stmt_, i, operator[](i));
		}
		void getData()
		{
			for (WORD i = 0; i < size(); ++i)
				ensure (SQL_SUCCEEDED(GetData[i](i)) || ODBC_ERROR(stmt_));
		}
	};
}
#endif // 0
