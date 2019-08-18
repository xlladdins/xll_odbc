// xllrange.h - range functions
#define EXCEL12
#pragma once
#include "../xll8/xll/xll.h"

typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xchar xchar;
typedef xll::traits<XLOPERX>::xword xword;

// "split,strings;;on,sep" -> ["split","strings","","on","sep"]
inline OPERX split(xcstr s, xword ns = 0, xcstr sep = _T(",;"))
{
	OPERX o;

	if (!ns)
		ns = _tcslen(s);

	xcstr b = s, e;
	do {
		e = _tcspbrk(b, sep);
		xchar n = e ? static_cast<xchar>(e - b) : ns;
		o.push_back(OPERX(b, n));
		ns -= n + 1;
		b = e + 1;
	} while (e);

	return o;
}

