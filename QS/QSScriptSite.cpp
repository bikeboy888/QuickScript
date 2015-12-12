// QSScriptSite.cpp : Implementation of CQSScriptSite

#include "stdafx.h"
#include "QSScriptSite.h"


// CQSScriptSite

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

STDMETHODIMP CQSScriptSite::put_ScriptEngine(BSTR bstrScriptEngine)
{
	HRESULT hr = S_OK;

    CScriptSite* pScriptSite = new CScriptSite();

    CComPtr<IActiveScript> spVBScript;
    CComPtr<IActiveScriptParse> spVBScriptParse;
    hr = spVBScript.CoCreateInstance(bstrScriptEngine);
    hr = spVBScript->SetScriptSite(pScriptSite);
    hr = spVBScript->QueryInterface(IID_IActiveScriptParse, (void**) &spVBScriptParse);
    hr = spVBScriptParse->InitNew();

    // Run some scripts
    CComVariant result;
    EXCEPINFO ei = { };
    hr = spVBScriptParse->ParseScriptText(OLESTR("1 + 2 + 3 + 4"), NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &result, &ei);
	CComVariant resultBSTR;
	hr = resultBSTR.ChangeType(VT_BSTR, &result);
	TCHAR szDebug[1024] = { };
	_stprintf(szDebug, _T("Result: %s\r\n"), (BSTR) V_BSTR(&resultBSTR));
	OutputDebugString(szDebug);

    // Cleanup
    spVBScriptParse = NULL;
    spVBScript = NULL;
    pScriptSite->Release();
    pScriptSite = NULL;

	return S_OK;
}

