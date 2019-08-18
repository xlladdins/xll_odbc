// xllrange.cpp - range operations
#include "xllrange.h"

#ifndef CATEGORY
#define CATEGORY _T("XLL")
#endif

using namespace xll;

typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xstring xstring;

static AddInX X_(xai_range_set)(
	FunctionX(XLL_HANDLEX, TX_("?xll_range_set"), _T("RANGE.SET"))
	.Range(_T("Range"), _T("is a range."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to Range."))
);
HANDLEX WINAPI X_(xll_range_set)(LPOPERX px)
{
#pragma XLLEXPORT
	return handle<OPERX>(new OPERX(*px)).get();
}

static AddInX X_(xai_range_get)(
	FunctionX(XLL_LPOPERX, TX_("?xll_range_get"), _T("RANGE.GET"))
	.Handle(_T("Handle"), _T("is a handle to a range."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the range correponding to Handle."))
);
LPOPERX WINAPI X_(xll_range_get)(HANDLEX h)
{
#pragma XLLEXPORT
	return handle<OPERX>(h, false).ptr();
}

static AddInX X_(xai_range_join)(
	FunctionX(XLL_LPOPERX, TX_("?xll_range_join"), _T("RANGE.JOIN"))
	.Range(_T("Range"), _T("is a range."))
	.Str(_T("FS"), _T("is the field separator."))
	.Str(_T("RS"), _T("is the record separator."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Join Range using field and record separators."))
);
LPOPERX WINAPI X_(xll_range_join)(const LPOPERX px, xcstr fs, xcstr rs)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX FS(fs);
		OPERX RS(rs);
		const OPERX& x(*px);

		r = x[0];
		for (xword i = 1; i < x.size(); ++i) {
			if (i%x.columns() == 0)
				r = XLL_XLF(Concatenate, r, RS);
			else
				r = XLL_XLF(Concatenate, r, FS);

			r = XLL_XLF(Concatenate, r, x[i]);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		r = OPERX(xlerr::NA);
	}

	return &r;
}

static AddInX X_(xai_range_split)(
	FunctionX(XLL_LPOPERX, TX_("?xll_range_split"), _T("RANGE.SPLIT"))
	.PStr(_T("String"), _T("is a string."))
	.Str(_T("FS"), _T("is a string of field separators."))
	.Str(_T("RS"), _T("is a string of record separators."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Split String using field and record separators."))
);
LPOPERX WINAPI X_(xll_range_split)(xcstr str, xcstr fs, xcstr rs)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		if (!*rs) {
			o = split(str + 1, str[0], fs);
		}
		else {
			OPERX r = split(str + 1, str[0], rs);
			o.resize(0,0);
			for (xword i = 0; i < r.size(); ++i) {
				OPERX f = split(r[i].val.str + 1, r[i].val.str[0], fs);
				o.push_back(f.resize(1, f.size()));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	return &o;
}