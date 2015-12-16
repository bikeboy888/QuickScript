// QSJScript.cpp : Implementation of CQSJScript

#include "stdafx.h"
#include "QSJScript.h"
#include "QSScriptSite.h"

// CQSJScript

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CQSJScript::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IQSJScript
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

HRESULT CQSJScript::FinalConstruct()
{
	HRESULT hr = S_OK;

	CHECKHR(CQSScriptSite::_CreatorClass::CreateInstance(NULL, IID_IQSScriptSite, (void**) &m_spIQSScriptSite));
	CHECKHR(m_spIQSScriptSite->put_ScriptEngine(CComBSTR(L"JScript")));

	CComBSTR bstrScript(
		L"function f(x) { return x + 1000; }\n"
		L"function jsEncodeURIComponent(uri) { return encodeURIComponent(uri); }\n"
		);
	CHECKHR(m_spIQSScriptSite->Execute(bstrScript, CComVariant()));

	CComVariant varResult;
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"f"), CComVariant(47), CComVariant(), CComVariant(), &varResult));

	varResult.Clear();
	CHECKHR(m_spIQSScriptSite->InvokeMethod(CComBSTR(L"jsEncodeURIComponent"), CComVariant(L"http://www.google.com?q=abc"), CComVariant(), CComVariant(), &varResult));

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
