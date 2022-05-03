// odbc.h - platform independent ODBC code
#include "xll/xll/ensure.h"
#include <Windows.h>
#include <sqlext.h>

typedef bool BIT;
typedef signed char STINYINT;
typedef unsigned char UTINYINT;
typedef __int64 SBIGINT; 
typedef unsigned __int64 UBIGINT;
typedef unsigned char* BINARY;

// https://docs.microsoft.com/en-us/sql/odbc/reference/appendixes/c-data-types
// C type identifier, ODBC C typedef, C type
#define ODBC_C_DATA_TYPES(X) \
X(CHAR, SQLCHAR*, unsigned char*) \
X(WCHAR, SQLWCHAR*, wchar_t*) \
X(SSHORT, SQLSMALLINT, short int) \
X(USHORT, SQLUSMALLINT, unsigned short int) \
X(SLONG, SQLINTEGER, long int) \
X(ULONG, SQLUINTEGER, unsigned long int) \
X(FLOAT, SQLREAL, float) \
X(DOUBLE, SQLDOUBLE, double) \
X(BIT, SQLCHAR, unsigned char) \
X(STINYINT, SQLSCHAR, signed char) \
X(UTINYINT, SQLCHAR, unsigned char) \
X(SBIGINT, SQLBIGINT, __int64) \
X(UBIGINT, SQLUBIGINT, unsigned __int64) \
X(BINARY, SQLCHAR*, unsigned char*) \
//X(BOOKMARK, BOOKMARK, unsigned long int) \
//X(VARBOOKMARK, SQLCHAR*, unsigned char*) \
//X(All C interval data types, SQL_INTERVAL_STRUCT, See the C Interval Structure section, later in this appendix.) \

#define ODBC_C_DATETIME_TYPES(X) \
X(TYPE_DATE, SQL_DATE_STRUCT, DATE_STRUCT) \
X(TYPE_TIME, SQL_TIME_STRUCT, TIME_STRUCT) \
X(TYPE_TIMESTAMP, SQL_TIMESTAMP_STRUCT, TIMESTAMP_STRUCT) \

// SQL_C_NUMERIC, SQL_C_GUID, 

namespace ODBC {

	// SQL_C_TYPE::CHAR = SQL_C_TYPE_CHAR, ...
	enum struct SQL_C_TYPE {
#define X_(a,b,c) a = SQL_C_##a,
		ODBC_C_DATA_TYPES(X_)
		//ODBC_C_DATETIME_TYPES(X_)
#undef X_
	};

	template<class T>
	inline T* Ptr(T& t)
	{
		return &t;
	}
	template<class T>
	inline SQLLEN Len(T& t)
	{
		return sizeof(t);
	}
	// counted string
	template<class SQLTCHAR>
	inline SQLTCHAR* Ptr(SQLTCHAR* str)
	{
		return str + 1;
	}
	template<class SQLTCHAR>
	inline SQLLEN Len(SQLTCHAR* str)
	{
		return str[0];
	}

	template<class T> struct traits {};
#define X_(a,b,c) template<> struct traits<a> { using a = b; static const SQLSMALLINT type = SQL_C_##a; };
	ODBC_C_DATA_TYPES(X_)
#undef X_

	// pointer to U sets t in destructor
	template<class U>
	class LenTPtr {
		SQLTCHAR* t;
		U len;
	public:
		LenTPtr(SQLTCHAR* t)
			: t(t), len(0)
		{ }
		~LenTPtr()
		{
			//if (len > std::numeric_limits<SQLTCHAR>::max()) {
			//	throw std::runtime_error(__FUNCTION__ ": length too long");
			//}

			t[0] = static_cast<SQLTCHAR>(len);
		}
		operator U* ()
		{
			return &len;
		}
	};
	template<class U, class T>
	inline auto LenPtr(T*)
	{
		return nullptr;
	}
	template<class U>
	inline auto LenPtr(SQLTCHAR* t)
	{
		return LenTPtr<U>(t);
	}

#define ODBC_STR(o) o.val.str + 1, o.val.str[0]
#define ODBC_BUF(o) ODBC_STR(o), ODBC::LenPtr<SQLSMALLINT>(o.val.str)
#define ODBC_BUF_(T, o) ODBC_STR(o), ODBC::LenPtr<T>(o.val.str)

	struct DiagRec {
		SQLTCHAR state[6] = { 0 };
		SQLTCHAR message[SQL_MAX_MESSAGE_LENGTH] = { 0 };
		SQLSMALLINT len;
		SQLINTEGER error;
		std::basic_string<SQLTCHAR> to_string() const
		{
			std::basic_string<SQLTCHAR> s(state);
			s.append({ ':', ' ' });
			s.append(message, len);

			return s;
		}
		SQLRETURN Get(SQLSMALLINT type, SQLHANDLE h, SQLSMALLINT n)
		{
			return SQLGetDiagRec(type, h, n, state, &error, message, SQL_MAX_MESSAGE_LENGTH, &len);
		}
		std::basic_string<SQLTCHAR> Get(SQLSMALLINT type, SQLHANDLE h)
		{
			std::basic_string<SQLTCHAR> s;

			for (SQLSMALLINT i = 1; SQL_SUCCEEDED(Get(type, h, i)); ++i) {
				s.append(to_string());
				s.append({ '\n' });
			}

			return s;
		}
	};

	enum class SQL_HANDLE {
		DBC = SQL_HANDLE_DBC,
		DESC = SQL_HANDLE_DESC,
		ENV = SQL_HANDLE_ENV,
		STMT = SQL_HANDLE_STMT,
	};

	// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlallochandle-function
	template<enum SQL_HANDLE T>
	struct Handle {
		SQLSMALLINT type() const
		{
			return static_cast<SQLSMALLINT>(T);
		}
		SQLRETURN rc;
		SQLHANDLE h;
		Handle(const SQLHANDLE& _h = SQL_NULL_HANDLE)
			: rc(SQLAllocHandle(type(), _h, &h))
		{ 
		}
		Handle(const Handle&) = delete;
		Handle& operator=(const Handle&) = delete;
		~Handle()
		{
			if (h)
				SQLFreeHandle(type(), h);
		}
		operator SQLHANDLE&()
		{
			return h;
		}
		operator const SQLHANDLE&() const
		{
			return h;
		}

		// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdiagrec-function
		std::basic_string<SQLTCHAR> GetDiagRec()
		{
			DiagRec rec;

			return rec.Get(type(), *this);
		}
		
		// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdiagfield-function
		SQLRETURN GetDiagField(
			SQLSMALLINT     RecNumber,
			SQLSMALLINT     DiagIdentifier,
			SQLPOINTER      DiagInfoPtr,
			SQLSMALLINT     BufferLength,
			SQLSMALLINT*    StringLengthPtr)
		{
			return SQLGetDiagField(type(), *this, RecNumber, DiagIdentifier, DiagInfoPtr, BufferLength, StringLengthPtr);
		}
	};

	// Environment singleton
	class Env {
		static SQLHANDLE env_(void)
		{
			static bool first(true);
			static Handle<SQL_HANDLE::ENV> h;

			if (first) {
				ensure (SQL_SUCCEEDED(SQLSetEnvAttr(h.h, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0)));
				first = false;
			}

			return h.h;
		}
	public:
		operator SQLHANDLE()
		{
			return env_();
		}
	};

	class Dbc : public Handle<SQL_HANDLE::DBC> {
        SQLTCHAR connect_[1024] = { 0 };
	public:
		Dbc()
            : Handle<SQL_HANDLE::DBC>(Env())
		{ }
        Dbc(const Dbc&) = delete;
        Dbc& operator=(const Dbc&) = delete;
        ~Dbc()
        { }
#ifdef _WINDOWS
		SQLRETURN DriverConnect(const SQLTCHAR* connect, SQLUSMALLINT complete = SQL_DRIVER_COMPLETE)
		{
			return rc = SQLDriverConnect(*this, GetDesktopWindow(), const_cast<SQLTCHAR*>(connect), SQL_NTS, connect_, 1024, 0, complete);
		}
#endif
		SQLRETURN BrowseConnect(const SQLTCHAR* connect)
		{
			return rc = SQLBrowseConnect(*this, const_cast<SQLTCHAR*>(connect), SQL_NTS, connect_, 1024, 0);
		}
		const SQLTCHAR* connectionString() const
		{
			return connect_;
		}
		SQLSMALLINT GetInfo(SQLSMALLINT info_type)
		{
			SQLSMALLINT si;

			ensure (SQL_SUCCEEDED(SQLGetInfo(*this, info_type, &si, SQL_IS_SMALLINT, 0)));

			return si;
		}

	};

	class Stmt : public Handle<SQL_HANDLE::STMT> {
	public:
		Stmt(const Dbc& dbc)
			: Handle<SQL_HANDLE::STMT>(dbc)
		{ }
		Stmt(const Stmt&) = delete;
		Stmt& operator=(const Stmt&) = delete;
		~Stmt()
		{
			SQLFreeStmt(*this, SQL_CLOSE);
		}

		SQLRETURN NumResultsCols(SQLSMALLINT& n) const
		{
			return SQLNumResultCols(*this, &n);
		}

		struct Col {
			SQLUSMALLINT ColumnNumber;
			SQLTCHAR     ColumnName[255] = {254};
			SQLSMALLINT  DataType;
			SQLULEN      ColumnSize;
			SQLSMALLINT  DecimalDigits;
			SQLSMALLINT  Nullable;
			// counted string
			const SQLTCHAR* Name() const
			{
				return ColumnName;
			}
			SQLSMALLINT Type() const
			{
				return DataType;
			}
			SQLULEN Size() const
			{
				return ColumnSize;
			}
			//SQLSMALLINT DecimalDigits;
			bool IsNullable() const
			{
				return SQL_NULLABLE == Nullable;
			}
		};
		SQLRETURN DescribeCol(
			SQLUSMALLINT ColumnNumber,
			SQLTCHAR*    ColumnName,
			SQLSMALLINT  BufferLength,
			SQLSMALLINT* NameLengthPtr,
			SQLSMALLINT* DataTypePtr,
			SQLULEN*     ColumnSizePtr,
			SQLSMALLINT* DecimalDigitsPtr,
			SQLSMALLINT* NullablePtr) const
		{
			return SQLDescribeCol(*this, ColumnNumber, ColumnName, BufferLength, NameLengthPtr,
				DataTypePtr, ColumnSizePtr, DecimalDigitsPtr, NullablePtr);
		}
		SQLRETURN DescribeCol(SQLUSMALLINT n, Col& col)
		{
			SQLSMALLINT len = 254;
			
			SQLRETURN ret = DescribeCol(n, col.ColumnName + 1, len, &len,
				&col.DataType, &col.ColumnSize, &col.DecimalDigits, &col.Nullable);
			col.ColumnName[0] = len;

			return ret;
		}

		// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlsetstmtattr-function
		SQLRETURN SetAttr(
			SQLINTEGER Attribute,
			SQLPOINTER ValuePtr,
			SQLINTEGER StringLength) const
		{
			return SQLSetStmtAttr(*this, Attribute, ValuePtr, StringLength);
		}
		template<class T>
		SQLRETURN SetAttr(SQLINTEGER Attribute, const T& t) const
		{
			return SQLSetStmtAttr(*this, Attribute, Ptr(t), Len(t));
		}

		// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlbindparameter-function
		SQLRETURN BindParameter(
			SQLUSMALLINT    ParameterNumber,
			SQLSMALLINT     InputOutputType,
			SQLSMALLINT     ValueType,
			SQLSMALLINT     ParameterType,
			SQLULEN         ColumnSize,
			SQLSMALLINT     DecimalDigits,
			SQLPOINTER      ParameterValuePtr,
			SQLLEN          BufferLength,
			SQLLEN*         StrLen_or_IndPtr)
		{
			return SQLBindParameter(*this, ParameterNumber, InputOutputType, ValueType, 
				ParameterType, ColumnSize, DecimalDigits, ParameterValuePtr, BufferLength, StrLen_or_IndPtr);
		}

		// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdata-function
		SQLRETURN GetData(
			SQLUSMALLINT   Col_or_Param_Num,
			SQLSMALLINT    TargetType,
			SQLPOINTER     TargetValuePtr,
			SQLLEN         BufferLength,
			SQLLEN*        StrLen_or_IndPtr)
		{
			return SQLGetData(*this, Col_or_Param_Num, TargetType, TargetValuePtr, BufferLength, StrLen_or_IndPtr);
		}
		template<class T>
		SQLRETURN GetData(SQLUSMALLINT n, T& t)
		{
			return GetData(n, traits<T>::type, Ptr(t), Len(t), LenPtr<SQLLEN>(t));
		}
	};

}