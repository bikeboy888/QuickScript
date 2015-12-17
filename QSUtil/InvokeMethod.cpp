#include "stdafx.h"
#include "InvokeMethod.h"

HRESULT InvokeMethod(IDispatch* pIDispatch, LPOLESTR szName, VARIANT varArg1, VARIANT varArg2, VARIANT varArg3, VARIANT* pvarResult)
{
	HRESULT hr = S_OK;

	if (pvarResult) VariantInit(pvarResult);

	DISPID DispID = DISPID_UNKNOWN;
	LCID lcid = 0;

	if (szName == NULL)
	{
		DispID = DISPID_VALUE;
	}
	else
	{
		hr = pIDispatch->GetIDsOfNames(IID_NULL, &szName, 1, lcid, &DispID);
		if (FAILED(hr)) return S_FALSE;
	}

	VARIANT* rgArgs[3] =
	{
		&varArg1,
		&varArg2,
		&varArg3
	};
	int nArgs = 0;
	for (int i = 0; i < 3; i++)
	{
		if (rgArgs[i]->vt != VT_EMPTY)
			nArgs = i + 1;
	}

	CComVariant Args[3];
	for (int i = 0; i < nArgs; i++)
	{
		Args[i] = *rgArgs[nArgs - 1 - i];
	}

	DISPPARAMS DispParams = { &Args[0], 0, nArgs, 0 };
	CComVariant varResult;
	hr = pIDispatch->Invoke(DispID, IID_NULL, lcid, DISPATCH_METHOD, &DispParams, &varResult, NULL, NULL);
	if (FAILED(hr)) return S_FALSE;

	if (pvarResult)
	{
		hr = varResult.Detach(pvarResult);
	}

	return S_OK;
}
