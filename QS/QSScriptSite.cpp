// QSScriptSite.cpp : Implementation of CQSScriptSite

#include "stdafx.h"
#include "QSScriptSite.h"

// CQSScriptSite

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSScriptSite
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

STDMETHODIMP CQSScriptSite::put_ScriptEngine(BSTR bstrScriptEngine)
{
	HRESULT hr = S_OK;

	if (m_spIActiveScript || m_spIActiveScriptParse || m_pScriptSite)
	{
		CHECKHR(Close());
	}

    m_pScriptSite = new CScriptSite();
    CHECKHR(m_spIActiveScript.CoCreateInstance(bstrScriptEngine));
    CHECKHR(m_spIActiveScript->SetScriptSite(m_pScriptSite));
    CHECKHR(m_spIActiveScript->QueryInterface(IID_IActiveScriptParse, (void**) &m_spIActiveScriptParse));
    CHECKHR(m_spIActiveScriptParse->InitNew());
	return S_OK;
}
//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::Close()
{
	m_spIActiveScript = NULL;
	m_spIActiveScriptParse = NULL;
	if (m_pScriptSite)
	{
		m_pScriptSite->Release();
		m_pScriptSite = NULL;
	}
	CoFreeUnusedLibrariesEx(0, 0);
	CoFreeUnusedLibrariesEx(0, 0);
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::Evaluate(BSTR bstrScript, VARIANT* pvarResult)
{
	HRESULT hr = S_OK;
	if (!pvarResult) return E_INVALIDARG;
	VariantInit(pvarResult);
	if (!m_spIActiveScriptParse) return E_POINTER;
    CComVariant result;
    EXCEPINFO ei = { };
    hr = m_spIActiveScriptParse->ParseScriptText(bstrScript, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &result, &ei);
	if (FAILED(hr)) return S_FALSE;
	CHECKHR(result.Detach(pvarResult));
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
