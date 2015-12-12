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

STDMETHODIMP CScriptSite::OnScriptError(IActiveScriptError *pIActiveScriptError)
{
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------
