// QSGlobal.cpp : Implementation of CQSGlobal

#include "stdafx.h"
#include "QSGlobal.h"
#include "DispatchHack.h"

// CQSGlobal

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSGlobal::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSGlobal
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

STDMETHODIMP CQSGlobal::GetRef(BSTR bstrName, IDispatch** ppIDispatch)
{
	HRESULT hr = S_OK;
	if (!ppIDispatch) return E_INVALIDARG;
	*ppIDispatch = NULL;
	if (!m_spIActiveScript) return S_FALSE;
	CComPtr<IDispatch> spIDispatch;
	CHECKHR(m_spIActiveScript->GetScriptDispatch(NULL, &spIDispatch));
	*ppIDispatch = (IDispatch*) new CDispatchHack(spIDispatch, bstrName, m_hWnd);
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSGlobal::Read(LPCOLESTR pszPropName, VARIANT *pVar, IErrorLog *pErrorLog)
{
	HRESULT hr = S_OK;
	if (!pVar) return E_INVALIDARG;
	VariantInit(pVar);
	if (wcsicmp(pszPropName, L"hWnd") == 0)
	{
		pVar->vt = VT_I4;
		V_I4(pVar) = (LONG) m_hWnd;
		return S_OK;
	}
	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSGlobal::Write(LPCOLESTR pszPropName, VARIANT *pVar)
{
	HRESULT hr = S_OK;
	if (wcsicmp(pszPropName, L"hWnd") == 0)
	{
		CComVariant varI4;
		CHECKHR(varI4.ChangeType(VT_I4, pVar));
		m_hWnd = (HWND) V_I4(&varI4);
		return S_OK;
	}
	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

