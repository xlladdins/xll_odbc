// xllrange.cpp - range operations
#include "xllrange.h"

#ifndef CATEGORY
#define CATEGORY "XL"
#endif

using namespace xll;

static AddIn xai_range_set(
	Function(XLL_HANDLE, "?xll_range_set", "RANGE.SET")
	.Arguments({
		Arg(XLL_LPOPER, "Range", "is a range.")
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return a handle to Range.")
);
HANDLEX WINAPI xll_range_set(LPOPER px)
{
#pragma XLLEXPORT
	return handle<OPER>(new OPER(*px)).get();
}

static AddIn xai_range_get(
	Function(XLL_LPOPER, "?xll_range_get", "RANGE.GET")
	.Arguments({
		Arg(XLL_HANDLE, "Handle", "is a handle to a range."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return the range correponding to Handle.")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	return handle<OPER>(h).ptr();
}

static AddIn xai_range_join(
	Function(XLL_LPOPER, "?xll_range_join", "RANGE.JOIN")
	.Arguments({
		Arg(XLL_LPOPER, "Range", "is a range."),
		Arg(XLL_CSTRING, "FS", "is the field separator."),
		Arg(XLL_CSTRING, "RS", "is the record separator."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Join Range using field and record separators.")
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

		r = OPER(ErrNA);
	}

	return &r;
}

static AddIn xai_range_split(
	Function(XLL_LPOPER, "?xll_range_split", "RANGE.SPLIT")
	.Arguments({
		Arg(XLL_PSTRING, "String", "is a string."),
		Arg(XLL_CSTRING, "FS", "is a string of field separators."),
		Arg(XLL_CSTRING, "RS", "is a string of record separators."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Split String using field and record separators.")
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

		o = OPER(ErrNA);
	}

	return &o;
}