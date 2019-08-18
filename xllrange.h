// xllrange.h - range functions
#define EXCEL12
#pragma once
#include <string.h>
#include "xll12/xll/xll.h"

// "split,strings;;on,sep" -> ["split","strings","","on","sep"]
inline xll::OPER split(const wchar_t* s, size_t ns = 0, const wchar_t* sep = L",;")
{
	xll::OPER o;

	if (!ns)
		ns = wcslen(s);

    const wchar_t* b = s;
    const wchar_t* e;
	do {
		e = wcspbrk(b, sep);
		wchar_t n = e ? static_cast<wchar_t>(e - b) : static_cast<wchar_t>(ns);
		o.push_back(xll::OPER(b, n));
		ns -= n + 1;
		b = e + 1;
	} while (e);

	return o;
}

