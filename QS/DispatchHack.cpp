#include "stdafx.h"
#include "DispatchHack.h"

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::QueryInterface(REFIID riid, void **ppvObject)
{
	if (!ppvObject) return E_POINTER;

	if (riid == IID_IUnknown || riid == IID_IDispatch)
	{
		AddRef();
		*ppvObject = (IDispatch*) this;
		return S_OK;
	}

	if (riid == IID_IPropertyBag)
	{
		AddRef();
		*ppvObject = (IPropertyBag*) this;
		return S_OK;
	}

	return E_NOINTERFACE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP_(ULONG) CDispatchHack::AddRef(void)
{
	return InterlockedIncrement(&m_RefCount);
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP_(ULONG) CDispatchHack::Release(void)
{
	if (InterlockedDecrement(&m_RefCount) == 0)
	{
		delete this;
		return 0;
	}

	return m_RefCount;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::GetTypeInfoCount(UINT *pctinfo)
{
	if (!pctinfo) return E_POINTER;
	*pctinfo = 0;
	return S_OK;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	if (ppTInfo == NULL)
		return E_INVALIDARG;
	*ppTInfo = NULL;

	if (iTInfo != 0)
		return DISP_E_BADINDEX;

	return S_FALSE;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT hr = S_OK;
	for (UINT i = 0; i < cNames; i++)
	{
		*rgDispId++ = DISPID_UNKNOWN;
		hr = DISP_E_UNKNOWNNAME;
	}
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr = S_OK;

	if (dispIdMember != DISPID_VALUE) return E_INVALIDARG;
	if (!m_spIDispatch) return E_POINTER;

	DISPID DispID = DISPID_UNKNOWN;

	hr = m_spIDispatch->GetIDsOfNames(IID_NULL, &((BSTR&) m_bstrMember), 1, lcid, &DispID);
	if (FAILED(hr)) return S_FALSE;

	hr = m_spIDispatch->Invoke(DispID, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	return hr;
}

//----------------------------------------------------------------------
//
//----------------------------------------------------------------------

STDMETHODIMP CDispatchHack::Read(LPCOLESTR pszPropName, VARIANT *pVar, IErrorLog *pErrorLog)
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

STDMETHODIMP CDispatchHack::Write(LPCOLESTR pszPropName, VARIANT *pVar)
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
