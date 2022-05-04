// odbc.h - platform independent ODBC code
#include "xll/xll/ensure.h"
#include <vector>
#define NOMINMAX
//#define WIN32_LEAN_AND_MEAN
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

#define SQL_DIAG_HEADER_FIELDS(X) \
X(CURSOR_ROW_COUNT, SQLLEN, "This field contains the count of rows in the cursor. Its semantics depend on the SQLGetInfo information types SQL_DYNAMIC_CURSOR_ATTRIBUTES2, SQL_FORWARD_ONLY_CURSOR_ATTRIBUTES2, SQL_KEYSET_CURSOR_ATTRIBUTES2, and SQL_STATIC_CURSOR_ATTRIBUTES2, which indicate which row counts are available for each cursor type (in the SQL_CA2_CRC_EXACT and SQL_CA2_CRC_APPROXIMATE bits).\nThe contents of this field are defined only for statement handles and only after SQLExecute, SQLExecDirect, or SQLMoreResults has been called. Calling SQLGetDiagField with a DiagIdentifier of SQL_DIAG_CURSOR_ROW_COUNT on other than a statement handle will return SQL_ERROR.") \
X(DYNAMIC_FUNCTION, SQLCHAR*, "This is a string that describes the SQL statement that the underlying function executed. (See "Values of the Dynamic Function fields, " later in this section, for specific values.) The contents of this field are defined only for statement handles and only after a call to SQLExecute, SQLExecDirect, or SQLMoreResults. Calling SQLGetDiagField with a DiagIdentifier of SQL_DIAG_DYNAMIC_FUNCTION on other than a statement handle will return SQL_ERROR. The value of this field is undefined before a call to SQLExecute or SQLExecDirect.") \
X(DYNAMIC_FUNCTION_CODE, SQLINTEGER, "This is a numeric code that describes the SQL statement that was executed by the underlying function. (See "Values of the Dynamic Function Fields, " later in this section, for specific value.) The contents of this field are defined only for statement handles and only after a call to SQLExecute, SQLExecDirect, or SQLMoreResults. Calling SQLGetDiagField with a DiagIdentifier of SQL_DIAG_DYNAMIC_FUNCTION_CODE on other than a statement handle will return SQL_ERROR. The value of this field is undefined before a call to SQLExecute or SQLExecDirect.") \
X(NUMBER, SQLINTEGER, "The number of status records that are available for the specified handle.") \
X(RETURNCODE, SQLRETURN, "Return code returned by the function. For a list of return codes, see Return Codes. The driver does not have to implement SQL_DIAG_RETURNCODE; it is always implemented by the Driver Manager. If no function has yet been called on the Handle, SQL_SUCCESS will be returned for SQL_DIAG_RETURNCODE.") \
X(ROW_COUNT, SQLLEN, "The number of rows affected by an insert, delete, or update performed by SQLExecute, SQLExecDirect, SQLBulkOperations, or SQLSetPos. It is driver-defined after a cursor specification has been executed. The contents of this field are defined only for statement handles. Calling SQLGetDiagField with a DiagIdentifier of SQL_DIAG_ROW_COUNT on other than a statement handle will return SQL_ERROR. The data in this field is also returned in the RowCountPtr argument of SQLRowCount. The data in this field is reset after every nondiagnostic function call, whereas the row count returned by SQLRowCount remains the same until the statement is set back to the prepared or allocated state.") \

#define SQL_DIAG_RECORD_FIELDS(X) \
X(CLASS_ORIGIN, SQLCHAR*, "A string that indicates the document that defines the class portion of the SQLSTATE value in this record. Its value is "ISO 9075" for all SQLSTATEs defined by Open Group and ISO call-level interface. For ODBC-specific SQLSTATEs (all those whose SQLSTATE class is "IM"), its value is "ODBC 3.0".") \
X(COLUMN_NUMBER, SQLINTEGER, "If the SQL_DIAG_ROW_NUMBER field is a valid row number in a rowset or a set of parameters, this field contains the value that represents the column number in the result set or the parameter number in the set of parameters. Result set column numbers always start at 1; if this status record pertains to a bookmark column, the field can be zero. Parameter numbers start at 1. It has the value SQL_NO_COLUMN_NUMBER if the status record is not associated with a column number or parameter number. If the driver cannot determine the column number or parameter number that this record is associated with, this field has the value SQL_COLUMN_NUMBER_UNKNOWN.\nThe contents of this field are defined only for statement handles.") \
X(CONNECTION_NAME, SQLCHAR*, "A string that indicates the name of the connection that the diagnostic record relates to. This field is driver-defined. For diagnostic data structures associated with the environment handle and for diagnostics that do not relate to any connection, this field is a zero-length string.") \
X(MESSAGE_TEXT, SQLCHAR*, "An informational message on the error or warning. This field is formatted as described in Diagnostic Messages. There is no maximum length to the diagnostic message text.") \
X(NATIVE, SQLINTEGER, "A driver/data source-specific native error code. If there is no native error code, the driver returns 0.") \
X(ROW_NUMBER, SQLLEN, "This field contains the row number in the rowset, or the parameter number in the set of parameters, with which the status record is associated. Row numbers and parameter numbers start with 1. This field has the value SQL_NO_ROW_NUMBER if this status record is not associated with a row number or parameter number. If the driver cannot determine the row number or parameter number that this record is associated with, this field has the value SQL_ROW_NUMBER_UNKNOWN.\nThe contents of this field are defined only for statement handles.") \
X(SERVER_NAME, SQLCHAR*, "A string that indicates the server name that the diagnostic record relates to. It is the same as the value returned for a call to SQLGetInfo with the SQL_DATA_SOURCE_NAME option. For diagnostic data structures associated with the environment handle and for diagnostics that do not relate to any server, this field is a zero-length string.") \
X(SUBCLASS_ORIGIN, SQLCHAR*, "A string with the same format and valid values as SQL_DIAG_CLASS_ORIGIN, that identifies the defining portion of the subclass portion of the SQLSTATE code. The ODBC-specific SQLSTATES for which "ODBC 3.0" is returned include the following:") \
//X(SQLSTATE, SQLCHAR*, "A five-character SQLSTATE diagnostic code. For more information, see SQLSTATEs.") \

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
	// counted string
	template<class SQLTCHAR>
	inline SQLTCHAR* Ptr(SQLTCHAR* str)
	{
		return str + 1;
	}

	template<class T>
	inline SQLLEN Len(T& t)
	{
		return sizeof(t);
	}
	// counted string
	template<class SQLTCHAR>
	inline SQLLEN Len(SQLTCHAR* str)
	{
		return str[0];
	}

	template<class T> struct traits {};
#define X_(a,b,c) template<> struct traits<a> { using a = b; static const SQLSMALLINT type = SQL_C_##a; };
	ODBC_C_DATA_TYPES(X_)
#undef X_

	// pointer to U sets t[0] in destructor
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
			t[0] = len > std::numeric_limits<SQLTCHAR>::max() ? 0 : static_cast<SQLTCHAR>(len);
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
#define ODBC_STR_BUF(o) ODBC_STR(o), ODBC::LenPtr<SQLSMALLINT>(o.val.str)
#define ODBC_STR_BUF_(T, o) ODBC_STR(o), ODBC::LenPtr<T>(o.val.str)

#define ODBC_PTR_LEN(x) ODBC::Ptr(x), ODBC::Len(x)
#define ODBC_PTR_LEN_BUF(U, x) ODBC_PTR_LEN(x), ODBC::LenPtr<U>(x)

	// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdiagrec-function
	struct DiagRec {
		SQLTCHAR state[6] = { 0 };
		SQLTCHAR message[SQL_MAX_MESSAGE_LENGTH] = { 0 };
		SQLSMALLINT len;
		SQLINTEGER error;
		SQLSMALLINT type;
		SQLHANDLE h;

		DiagRec(SQLSMALLINT type, SQLHANDLE h)
			: type(type), h(h)
		{ }
		DiagRec(DiagRec&) = delete;
		DiagRec& operator=(DiagRec&) = delete;
		~DiagRec()
		{ }
		std::basic_string<SQLTCHAR> to_string() const
		{
			std::basic_string<SQLTCHAR> s(state);

			s.append({ ':', ' ' });
			s.append(message, len);

			return s;
		}
		// Get one record
		SQLRETURN Get(SQLSMALLINT n)
		{
			return SQLGetDiagRec(type, h, n, state, &error, message, SQL_MAX_MESSAGE_LENGTH, &len);
		}
		// Get all records
		std::vector<std::basic_string<SQLTCHAR>> Get()
		{
			std::vector<std::basic_string<SQLTCHAR>> s;

			for (SQLSMALLINT i = 1; SQL_SUCCEEDED(Get(i)); ++i) {
				s.push_back(to_string());
			}

			return s;
		}
	};

	// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlgetdiagfield-function
	template<class U>
	inline SQLRETURN GetDiagField(SQLSMALLINT HandleType, SQLHANDLE Handle, SQLSMALLINT RecNumber, SQLSMALLINT DiagIdentifier, U& u)
	{
		return SQLGetDiagField(HandleType, Handle, RecNumber, DiagIdentifier, Ptr(u), Len(u), LenPtr<SQLSMALLINT>(u));
	}
	class DiagField {
		SQLSMALLINT type;
		SQLHANDLE h;
	public:
		DiagField(SQLSMALLINT type, SQLHANDLE h)
			: type(type), h(h)
		{ }
		DiagField(const DiagField&) = delete;
		DiagField& operator=(const DiagField&) = delete;
		~DiagField()
		{ }
		template<class U>
		U Get(SQLSMALLINT n, SQLSMALLINT id)
		{
			U u;

			GetDiagField<U>(type, h, n, id, u);

			return u;
		}
		/*
		// s points to counted, allocated memory
		SQLRETURN Get(SQLSMALLINT n, SQLSMALLINT id, SQLTCHAR* s)
		{
			return 0; // GetDiagField(type, h, n, id, s);
		}
		*/
		
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
		
	};

	template<enum SQL_HANDLE T>
	inline auto GetDiagRec(Handle<T>& h)
	{
		DiagRec rec(h.type(), h);

		return rec.Get();
	}

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
		using cstr = const SQLTCHAR*;
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