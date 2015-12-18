#include "stdafx.h"
#include "OutputDebugFormat.h"

void OutputDebugFormat(LPCWSTR pFormat, ...)
{
	va_list args;
	va_start(args, pFormat);
	WCHAR szText[1024] = { };
	vswprintf(szText, pFormat, args);
	va_end(args);
	OutputDebugString(szText);
}

