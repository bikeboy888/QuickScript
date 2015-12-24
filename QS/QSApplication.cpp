// QSApplication.cpp : Implementation of CQSApplication

#include "stdafx.h"
#include "QSApplication.h"


// CQSApplication

STDMETHODIMP CQSApplication::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSApplication
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSApplication::get_UserProperties(BSTR bstrName, VARIANT* pvarValue)
{
	HRESULT hr = S_OK;
	if (!pvarValue) return E_INVALIDARG;
	VariantInit(pvarValue);
	CComVariant varValue;
	if (!m_UserProperties.Lookup(bstrName, varValue))
	{
		return S_FALSE;
	}
	CHECKHR(varValue.Detach(pvarValue));
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSApplication::put_UserProperties(BSTR bstrName, VARIANT varValue)
{
	POSITION pos = m_UserProperties.SetAt(bstrName, varValue);
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
