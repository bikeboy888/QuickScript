// QSScriptSite.cpp : Implementation of CQSScriptSite

#include "stdafx.h"
#include "QSScriptSite.h"
#include "..\QSUtil\InvokeMethod.h"
#include "QSGlobal.h"

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

STDMETHODIMP CQSScriptSite::get_hWnd(OLE_HANDLE* phWnd)
{
	HRESULT hr = S_OK;
	if (!phWnd) return E_INVALIDARG;
	if (!m_pScriptSite) return E_POINTER;
	HWND hWnd = 0;
	CHECKHR(m_pScriptSite->GetWindow(&hWnd));
	*phWnd = (OLE_HANDLE) hWnd;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::put_hWnd(OLE_HANDLE hWnd)
{
	HRESULT hr = S_OK;
	if (!m_pScriptSite) return E_POINTER;
	CHECKHR(m_pScriptSite->SetWindow((HWND) hWnd));
	if (m_spIQSGlobal)
	{
		CComPtr<IPropertyBag> spIPropertyBag;
		hr = m_spIQSGlobal->QueryInterface(IID_IPropertyBag, (void**) &spIPropertyBag);
		if (SUCCEEDED(hr) && spIPropertyBag)
		{
			CComVariant varhWnd((LONG) hWnd);
			CHECKHR(spIPropertyBag->Write(OLESTR("hWnd"), &varhWnd));
		}
	}
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::put_ScriptEngine(BSTR bstrScriptEngine)
{
	HRESULT hr = S_OK;

	if (m_spIActiveScript || m_spIActiveScriptParse || m_pScriptSite || m_spIQSGlobal)
	{
		CHECKHR(Close());
	}

    CHECKHR(m_spIActiveScript.CoCreateInstance(bstrScriptEngine));
    m_pScriptSite = new CScriptSite();
    CHECKHR(m_spIActiveScript->SetScriptSite(m_pScriptSite));
    CHECKHR(m_spIActiveScript->QueryInterface(IID_IActiveScriptParse, (void**) &m_spIActiveScriptParse));
    CHECKHR(m_spIActiveScriptParse->InitNew());
    CHECKHR(m_spIActiveScript->SetScriptState(SCRIPTSTATE_CONNECTED));

	CComObject<CQSGlobal>* pQSGlobal = NULL;
	CHECKHR(CComObject<CQSGlobal>::CreateInstance(&pQSGlobal));
	pQSGlobal->m_spIActiveScript = m_spIActiveScript;
	CHECKHR(pQSGlobal->QueryInterface(IID_IQSGlobal, (void**) &m_spIQSGlobal));
	CHECKHR(SetItem(CComBSTR(L"Global"), SCRIPTITEM_ISVISIBLE | SCRIPTITEM_GLOBALMEMBERS, m_spIQSGlobal));

	return S_OK;
}
//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::Close()
{
	HRESULT hr = S_OK;
	if (m_spIQSGlobal)
	{
		m_spIQSGlobal = NULL;
	}
	if (m_spIActiveScript)
	{
		//hr = m_spIActiveScript->SetScriptState(SCRIPTSTATE_INITIALIZED);
		//hr = m_spIActiveScript->SetScriptState(SCRIPTSTATE_DISCONNECTED);
		hr = m_spIActiveScript->Close();
		m_spIActiveScript = NULL;
	}
	m_spIActiveScriptParse = NULL;
	if (m_pScriptSite)
	{
		m_pScriptSite->Close();
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

STDMETHODIMP CQSScriptSite::Evaluate(BSTR bstrScript, VARIANT varContext, VARIANT* pvarResult)
{
	HRESULT hr = S_OK;
	if (!pvarResult) return E_INVALIDARG;
	VariantInit(pvarResult);
	CHECKHR(ParseScriptText(bstrScript, varContext, SCRIPTTEXT_ISEXPRESSION, pvarResult));
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::Execute(BSTR bstrScript, VARIANT varContext)
{
	HRESULT hr = S_OK;
	//CHECKHR(ParseScriptText(bstrScript, varContext, SCRIPTTEXT_ISPERSISTENT, NULL));
	CHECKHR(ParseScriptText(bstrScript, varContext, 0, NULL));
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::InvokeMethod(BSTR bstrName, VARIANT varArg1, VARIANT varArg2, VARIANT varArg3, VARIANT* pvarResult)
{
	HRESULT hr = S_OK;
	if (pvarResult) VariantInit(pvarResult);
	if (!m_spIActiveScript) return S_FALSE;
	CComPtr<IDispatch> spIDispatch;
	CHECKHR(m_spIActiveScript->GetScriptDispatch(NULL, &spIDispatch));
	CHECKHR(::InvokeMethod(spIDispatch, (LPOLESTR) bstrName, varArg1, varArg2, varArg3, pvarResult));
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::ParseScriptText(BSTR bstrScript, VARIANT varContext, DWORD dwFlags, VARIANT* pvarResult)
{
	HRESULT hr = S_OK;
	if (pvarResult) VariantInit(pvarResult);
	if (!m_spIActiveScriptParse) return E_POINTER;
	DWORD dwContext = 0;
    CComVariant result;
    EXCEPINFO ei = { };
    hr = m_spIActiveScriptParse->ParseScriptText(bstrScript, NULL, NULL, NULL, 0, 0, dwFlags, &result, &ei);
	if (FAILED(hr)) return S_FALSE;
	if (pvarResult)
	{
		CHECKHR(result.Detach(pvarResult));
	}
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::SetItem(BSTR bstrName, LONG nFlags, IDispatch* pIDispatch)
{
	HRESULT hr = S_OK;
	if (!m_pScriptSite) return E_POINTER;
	if (!m_spIActiveScript) return E_POINTER;
	CHECKHR(m_pScriptSite->SetItemInfo(bstrName, pIDispatch));
	CHECKHR(m_spIActiveScript->AddNamedItem(bstrName, (DWORD) nFlags));
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
