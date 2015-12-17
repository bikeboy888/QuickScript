#include "stdafx.h"
#include "VariantToBSTR.h"

HRESULT VariantToBSTR(VARIANT& var, BSTR* pbstr)
{
	HRESULT hr = S_OK;
	if (!pbstr) return E_INVALIDARG;
	*pbstr = NULL;
	if (var.vt == VT_NULL || var.vt == VT_EMPTY) return S_OK;
	CComVariant varBSTR;
	CHECKHR(varBSTR.ChangeType(VT_BSTR, &var));
	VARIANT vBSTR = { };
	CHECKHR(varBSTR.Detach(&vBSTR));
	*pbstr = V_BSTR(&vBSTR);
	return S_OK;
}
