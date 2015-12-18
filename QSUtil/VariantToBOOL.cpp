#include "stdafx.h"
#include "VariantToBOOL.h"

HRESULT VariantToBOOL(VARIANT& var, VARIANT_BOOL* pbValue, VARIANT_BOOL bDefaultValue)
{
	HRESULT hr = S_OK;
	if (!pbValue) return E_INVALIDARG;
	*pbValue = bDefaultValue;
	VARIANT* pvar = &var;
	if (pvar->vt == (VT_VARIANT | VT_BYREF))
	{
		pvar = pvar->pvarVal;
	}
	if (pvar->vt == VT_EMPTY || pvar->vt == VT_NULL) return S_OK;
	CComVariant varBOOL;
	hr = varBOOL.ChangeType(VT_BOOL, &var);
	if (SUCCEEDED(hr))
	{
		*pbValue = V_BOOL(&varBOOL);
		return S_OK;
	}
	return hr;
}
