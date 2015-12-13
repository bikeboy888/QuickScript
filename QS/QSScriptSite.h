// QSScriptSite.h : Declaration of the CQSScriptSite

#pragma once
#ifdef STANDARDSHELL_UI_MODEL
#include "resource.h"
#endif
#ifdef POCKETPC2003_UI_MODEL
#include "resourceppc.h"
#endif
#ifdef SMARTPHONE2003_UI_MODEL
#include "resourcesp.h"
#endif
#ifdef AYGSHELL_UI_MODEL
#include "resourceayg.h"
#endif

#include "QS_i.h"
#include "ScriptSite.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CQSScriptSite

class ATL_NO_VTABLE CQSScriptSite :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CQSScriptSite, &CLSID_QSScriptSite>,
	public ISupportErrorInfo,
	public IDispatchImpl<IQSScriptSite, &IID_IQSScriptSite, &LIBID_QSLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CQSScriptSite() : m_pScriptSite(NULL)
	{
	}

#ifndef _CE_DCOM
DECLARE_REGISTRY_RESOURCEID(IDR_QSSCRIPTSITE)
#endif


BEGIN_COM_MAP(CQSScriptSite)
	COM_INTERFACE_ENTRY(IQSScriptSite)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
		if (m_spIActiveScript || m_spIActiveScriptParse || m_pScriptSite)
		{
			Close();
		}
	}

public:
	CScriptSite* m_pScriptSite;
	CComPtr<IActiveScript> m_spIActiveScript;
	CComPtr<IActiveScriptParse> m_spIActiveScriptParse;

public:
	STDMETHOD(put_ScriptEngine)(BSTR bstrScriptEngine);
	STDMETHOD(Close)();
	STDMETHOD(Evaluate)(BSTR bstrScript, VARIANT* pvarResult);

};

OBJECT_ENTRY_AUTO(__uuidof(QSScriptSite), CQSScriptSite)
