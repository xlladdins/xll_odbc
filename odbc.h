// odbc.h - platform independent ODBC code
#include "xll12/xll/ensure.h"
#include <Windows.h>
#include <sqlext.h>

namespace ODBC {

	template<SQLSMALLINT T>
	struct Handle {
		static const SQLSMALLINT type = T;
		SQLRETURN rc;
		SQLHANDLE h;
		Handle(const SQLHANDLE& _h = SQL_NULL_HANDLE)
			: rc(SQLAllocHandle(T, _h, &h))
		{ 
		}
		Handle(const Handle&) = delete;
		Handle& operator=(const Handle&) = delete;
		~Handle()
		{
			if (h)
				SQLFreeHandle(T, h);
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

	template<SQLSMALLINT T>
	class DiagRec {
		const Handle<T>& h_;
	public:
		SQLTCHAR state[6], message[SQL_MAX_MESSAGE_LENGTH]; 
		SQLINTEGER error;
		SQLRETURN rc;
		DiagRec(const Handle<T>& h)
			: h_(h), error(0), rc(SQL_SUCCESS)
		{ }
		DiagRec(const DiagRec&) = delete;
		DiagRec& operator=(const DiagRec&) = delete;
		~DiagRec()
		{ }
		SQLRETURN Get(SQLSMALLINT n)
		{
			SQLSMALLINT len;

			return rc = SQLGetDiagRec(T, h_, n, state, &error, message, SQL_MAX_MESSAGE_LENGTH, &len);
		}
	};

/*	// proxy class for L*
	template<class L, class T>
	struct LenPtr {
		L& len;
		LenPtr(L& l, T& t)
			: len(l), type(t)
		{ }
		~LenPtr()
		operator L*()
		{
			return &len;
		}
	};
	template<class L>
	struct LenPtr<L,SQLTCHAR*> {
		~LenPtr()
		{
			type[0] = static_cast<SQLTCHAR>(len);
		}
	};
*/
	template<class R>
	struct field_traits {
		static const SQLSMALLINT len;
	};
#define ODBC_FIELD_TRAITS(t) template<> struct field_traits<SQL ## t > { static const SQL ## t len = SQL_IS_ ## t ; };
	ODBC_FIELD_TRAITS(SMALLINT)
	ODBC_FIELD_TRAITS(USMALLINT)
	ODBC_FIELD_TRAITS(INTEGER)
	ODBC_FIELD_TRAITS(UINTEGER)

	template<SQLSMALLINT T>
	class DiagField {
		const Handle<T>& h_;
	public:
		DiagField(const Handle<T>& h)
			: h_(h)
		{ }
		~DiagField()
		{ }
		template<class R>
		R Get(SQLSMALLINT n, SQLSMALLINT id)
		{
			R r;

			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, &r, field_traits<R>::len(r), 0)));

			return r;
		}
		std::basic_string<SQLTCHAR> Get(SQLSMALLINT n, SQLSMALLINT id)
		{
			std::basic_string<SQLTCHAR> r(255);

			SQLSMALLINT len;
			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, &r[0], r.size(), &len)));
			if (len < r.size())
				r.resize(len);

			return r;
		}
		void Get(SQLSMALLINT n, SQLSMALLINT id, SQLTCHAR* str)
		{
			SQLSMALLINT len;
			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, str + 1, str[0], &len)));
			if (len < r.size())
				str[0] = static_cast<SQLTCHAR>(len);
		}
	};

	// Environment singleton
	class Env {
		static SQLHANDLE env_(void)
		{
			static bool first(true);
			static Handle<SQL_HANDLE_ENV> h;

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

	class Dbc : public Handle<SQL_HANDLE_DBC> {
		SQLTCHAR connect_[1024];
	public:
		Dbc()
            : Handle<SQL_HANDLE_DBC>(Env())
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
		SQLSMALLINT GetInfo(SQLSMALLINT type)
		{
			SQLSMALLINT si;

			ensure (SQL_SUCCEEDED(SQLGetInfo(*this, type, &si, SQL_IS_SMALLINT, 0)));

			return si;
		}

	};

	class Stmt : public Handle<SQL_HANDLE_STMT> {
	public:
		Stmt(const Dbc& dbc)
			: Handle<SQL_HANDLE_STMT>(dbc)
		{ }
        Stmt(const Stmt&) = delete;
        Stmt& operator=(const Stmt&) = delete;
        ~Stmt()
		{
			SQLFreeStmt(*this, SQL_CLOSE);
		}
	};

}