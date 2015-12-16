// QSScriptSite.cpp : Implementation of CQSScriptSite

#include "stdafx.h"
#include "QSScriptSite.h"
#include "..\QSUtil\InvokeMethod.h"

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

    CHECKHR(m_spIActiveScript.CoCreateInstance(bstrScriptEngine));
    m_pScriptSite = new CScriptSite();
    CHECKHR(m_spIActiveScript->SetScriptSite(m_pScriptSite));
    CHECKHR(m_spIActiveScript->QueryInterface(IID_IActiveScriptParse, (void**) &m_spIActiveScriptParse));
    CHECKHR(m_spIActiveScriptParse->InitNew());
    CHECKHR(m_spIActiveScript->SetScriptState(SCRIPTSTATE_CONNECTED));
	return S_OK;
}
//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::Close()
{
	HRESULT hr = S_OK;
	if (m_spIActiveScript)
	{
		hr = m_spIActiveScript->SetScriptState(SCRIPTSTATE_INITIALIZED);
		//hr = m_spIActiveScript->SetScriptState(SCRIPTSTATE_DISCONNECTED);
		hr = m_spIActiveScript->Close();
		m_spIActiveScript = NULL;
	}
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
	CHECKHR(ParseScriptText(bstrScript, varContext, SCRIPTTEXT_ISPERSISTENT, NULL));
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
	//CComPtr<IDispatch> spIDispatch;
	IDispatch* spIDispatch = NULL;
	CHECKHR(m_spIActiveScript->GetScriptDispatch(NULL, &spIDispatch));
	CHECKHR(::InvokeMethod(spIDispatch, (LPOLESTR) bstrName, varArg1, varArg2, varArg3, pvarResult));
	spIDispatch = NULL;
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSScriptSite::ImportScript(BSTR bstrScript, BSTR bstrScriptEngine, VARIANT varContext)
{
	HRESULT hr = S_OK;
	//CoInitializeEx(NULL, COINIT_MULTITHREADED);
	CComPtr<IQSScriptSite> spIQSScriptSite;
	CHECKHR(CQSScriptSite::_CreatorClass::CreateInstance(NULL, IID_IQSScriptSite, (void**) &spIQSScriptSite));
	CHECKHR(spIQSScriptSite->put_ScriptEngine(bstrScriptEngine));
	VARIANT varContextEmpty = { };
	CHECKHR(spIQSScriptSite->Execute(bstrScript, varContextEmpty));
	CComVariant varResult;
	CHECKHR(spIQSScriptSite->Evaluate(CComBSTR(L"f(5)"), varContextEmpty, &varResult));
	CHECKHR(spIQSScriptSite->Close());
	CComVariant varResultBSTR;
	CHECKHR(varResultBSTR.ChangeType(VT_BSTR, &varResult));
	TCHAR szText[1024] = { };
	_stprintf(szText, _T("Result: f(5) = %s\r\n"), V_BSTR(&varResultBSTR));
	OutputDebugString(szText);
	spIQSScriptSite = NULL;
	//CoUninitialize();
	::CoFreeUnusedLibrariesEx(0, 0);
	::CoFreeUnusedLibrariesEx(0, 0);
	return S_OK;
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
