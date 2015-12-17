#include "stdafx.h"

void OutputDebugVariant(VARIANT& var)
{
	switch (var.vt)
	{
	case VT_EMPTY:
		OutputDebugString(L"VT_EMPTY");
		return;

	case VT_NULL:
		OutputDebugString(L"VT_NULL");
		return;
	}

	CComVariant varBSTR;
	varBSTR.ChangeType(VT_BSTR, &var);
	OutputDebugString(V_BSTR(&varBSTR));

	OutputDebugString(L" ");

	TCHAR szText[1024] = { };
	switch (var.vt)
	{
	case VT_I4:			OutputDebugString(L"VT_I4");		break;
	case VT_R8:			OutputDebugString(L"VT_R8");		break;
	case VT_DISPATCH:	OutputDebugString(L"VT_DISPATCH");	break;
	case VT_BSTR:		OutputDebugString(L"VT_BSTR");		break;
	default:
		_stprintf(szText, _T("VarType:%d"), var.vt);
		OutputDebugString(szText);
	}
}
