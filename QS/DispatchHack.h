#pragma once

class CDispatchHack : public IDispatch, public IPropertyBag
{
public:
	CDispatchHack(IDispatch* pIDispatch = NULL, LPCWSTR bszMember = NULL, HWND hWnd = NULL) :
		m_RefCount(1),
		m_hWnd(hWnd),
		m_spIDispatch(pIDispatch),
		m_bstrMember(bszMember) { }

// IUnknown
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

// IDispatch
	STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
	STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

// IPropertyBag
	STDMETHOD(Read)(LPCOLESTR pszPropName, VARIANT *pVar, IErrorLog *pErrorLog);
	STDMETHOD(Write)(LPCOLESTR pszPropName, VARIANT *pVar);

protected:
	LONG m_RefCount;
	HWND m_hWnd;

public:
	CComPtr<IDispatch> m_spIDispatch;
	CComBSTR m_bstrMember;

};
