// xllrange.cpp - range operations
#include "xllrange.h"

#ifndef CATEGORY
#define CATEGORY L"XLL"
#endif

using namespace xll;

static AddIn xai_range_set(
	Function(XLL_HANDLE, L"?xll_range_set", L"RANGE.SET")
	.Arg(XLL_LPOPER, L"Range", L"is a range.")
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(L"Return a handle to Range.")
);
HANDLEX WINAPI xll_range_set(LPOPER px)
{
#pragma XLLEXPORT
	return handle<OPER>(new OPER(*px)).get();
}

static AddIn xai_range_get(
	Function(XLL_LPOPER, L"?xll_range_get", L"RANGE.GET")
	.Arg(XLL_HANDLE, L"Handle", L"is a handle to a range.")
	.Category(CATEGORY)
	.FunctionHelp(L"Return the range correponding to Handle.")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	return handle<OPER>(h).ptr();
}

static AddIn xai_range_join(
	Function(XLL_LPOPER, L"?xll_range_join", L"RANGE.JOIN")
	.Arg(XLL_LPOPER, L"Range", L"is a range.")
	.Arg(XLL_CSTRING, L"FS", L"is the field separator.")
    .Arg(XLL_CSTRING, L"RS", L"is the record separator.")
	.Category(CATEGORY)
	.FunctionHelp(L"Join Range using field and record separators.")
);
LPOPER WINAPI xll_range_join(const LPOPER px, const wchar_t* fs, const wchar_t* rs)
{
#pragma XLLEXPORT
	static OPER r;

	try {
		OPER FS(fs);
		OPER RS(rs);
		const OPER& x(*px);

		r = x[0];
		for (WORD i = 1; i < x.size(); ++i) {
			if (i%x.columns() == 0)
				r = r & RS;
			else
				r = r & FS;

			r = r & x[i];
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		r = OPER(xlerr::NA);
	}

	return &r;
}

static AddIn xai_range_split(
	Function(XLL_LPOPER, L"?xll_range_split", L"RANGE.SPLIT")
	.Arg(XLL_PSTRING, L"String", L"is a string.")
	.Arg(XLL_CSTRING, L"FS", L"is a string of field separators.")
	.Arg(XLL_CSTRING, L"RS", L"is a string of record separators.")
	.Category(CATEGORY)
	.FunctionHelp(L"Split String using field and record separators.")
);
LPOPER WINAPI xll_range_split(const wchar_t* str, const wchar_t* fs, const wchar_t* rs)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (!*rs) {
			o = split(str + 1, str[0], fs);
		}
		else {
			OPER r = split(str + 1, str[0], rs);
			o.resize(0,0);
			for (WORD i = 0; i < r.size(); ++i) {
				OPER f = split(r[i].val.str + 1, r[i].val.str[0], fs);
				o.push_back(f.resize(1, f.size()));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPER(xlerr::NA);
	}

	return &o;
}