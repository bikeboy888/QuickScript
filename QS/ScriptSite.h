#pragma once

class CScriptSite :
    public IActiveScriptSite,
    public IActiveScriptSiteWindow
{
public:
    CScriptSite() : m_cRefCount(1), m_hWnd(NULL) { }

    // IUnknown

    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject);

    // IActiveScriptSite

    STDMETHOD(GetLCID)(LCID *plcid){ *plcid = 0; return S_OK; }
    STDMETHOD(GetItemInfo)(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti);
    STDMETHOD(GetDocVersionString)(BSTR *pbstrVersion) { *pbstrVersion = SysAllocString(L"1.0"); return S_OK; }
    STDMETHOD(OnScriptTerminate)(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo) { return S_OK; }
    STDMETHOD(OnStateChange)(SCRIPTSTATE ssScriptState) { return S_OK; }
    STDMETHOD(OnScriptError)(IActiveScriptError *pIActiveScriptError);
    STDMETHOD(OnEnterScript)(void) { return S_OK; }
    STDMETHOD(OnLeaveScript)(void) { return S_OK; }

    // IActiveScriptSiteWindow

    STDMETHOD(GetWindow)(HWND *phWnd) { *phWnd = m_hWnd; return S_OK; }
    STDMETHOD(EnableModeless)(BOOL fEnable) { return S_OK; }

    // Miscellaneous

	STDMETHOD(Close)();
    STDMETHOD(SetWindow)(HWND hWnd) { m_hWnd = hWnd; return S_OK; }
	STDMETHOD(SetItemInfo)(LPCOLESTR szName, IUnknown* piunkItem);

protected:
	CAtlMap<CComBSTR, CComPtr<IUnknown>, CElementTraits<CComBSTR>> m_ItemInfo;

public:
    LONG m_cRefCount;
    HWND m_hWnd;

};
