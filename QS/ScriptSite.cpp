#include "stdafx.h"
#include "ScriptSite.h"

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP_(ULONG) CScriptSite::AddRef()
{
    return InterlockedIncrement(&m_cRefCount);
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP_(ULONG) CScriptSite::Release()
{
    if (!InterlockedDecrement(&m_cRefCount))
    {
        delete this;
        return 0;
    }
    return m_cRefCount;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
    if (riid == IID_IUnknown || riid == IID_IActiveScriptSiteWindow)
    {
        *ppvObject = (IActiveScriptSiteWindow *) this;
        AddRef();
        return NOERROR;
    }
    if (riid == IID_IActiveScriptSite)
    {
        *ppvObject = (IActiveScriptSite *) this;
        AddRef();
        return NOERROR;
    }
    return E_NOINTERFACE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

class CExcepInfo : public EXCEPINFO
{
public:
	~CExcepInfo() { Clear(); }

	void Clear()
	{
		if (bstrSource)
		{
			SysFreeString(bstrSource);
			bstrSource = NULL;
		}
		if (bstrDescription)
		{
			SysFreeString(bstrDescription);
			bstrDescription = NULL;
		}
		if (bstrHelpFile)
		{
			SysFreeString(bstrHelpFile);
			bstrHelpFile = NULL;
		}
	}

};

STDMETHODIMP CScriptSite::OnScriptError(IActiveScriptError *pIActiveScriptError)
{
	HRESULT hr = S_OK;
	if (!pIActiveScriptError) return S_OK;

	CExcepInfo ei;
	hr = pIActiveScriptError->GetExceptionInfo(&ei);

	CComBSTR bstrSource;
	hr = pIActiveScriptError->GetSourceLineText(&bstrSource);

	DWORD dwContext = 0;
	ULONG nLine = 0;
	LONG nPos = 0;
	hr = pIActiveScriptError->GetSourcePosition(&dwContext, &nLine, &nPos);

	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
